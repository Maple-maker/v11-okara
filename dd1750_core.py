"""DD1750 core - Fixed with smart description extraction."""

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


ROWS_PER_PAGE = 18
PAGE_W, PAGE_H = letter

X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP = 616.0
Y_TABLE_BOTTOM = 89.5
ROW_H = (Y_TABLE_TOP - Y_TABLE_BOTTOM) / ROWS_PER_PAGE
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def _extract_middle_text_smart(text: str) -> str:
    """
    Extract the middle/actual description from a table cell.
    
    Logic:
    1. Split text by newlines.
    2. Filter out garbage lines (too short, all caps, look like codes).
    3. Select the best line based on length and content.
    """
    if not text:
        return ""
    
    lines = [ln for ln in text.strip().split('\n') if ln.strip()]
    
    if not lines:
        return ""
    
    # Filter out obvious garbage
    candidates = []
    for ln in lines:
        # Skip single characters or very short strings
        if len(ln) < 5:
            continue
        
        # Skip strings that are all caps (likely headers/codes)
        if ln.isupper() and not any(c.islower() for c in ln):
            continue
        
        # Skip strings that look like codes (WTY, ARC, etc.) even if mixed case
        if re.match(r'^(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G|TMY|SMD|ECC)$', ln.upper()):
            continue
        
        candidates.append(ln)
    
    if not candidates:
        # If all filtered out, check first line again
        line_1 = lines[0].strip()
        if len(line_1) > 10:
            return line_1  # Use the first line if it's substantial
        return ""
    
    # Select best candidate
    # Prefer longer lines, lines with mixed case (more likely to be nomenclature)
    # Prefer lines with numbers or special characters
    candidates.sort(key=lambda x: (len(x), any(c.islower() for c in x)), reverse=True)
    
    description = candidates[0].strip()
    
    # Clean final description
    # Remove trailing garbage
    description = re.sub(r'\s+', ' ', description).strip()
    
    return description


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
                    lv_idx = desc_idx = mat_idx = auth_idx = -1
                    
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
                                auth_idx = i
                    
                    if lv_idx == -1 or desc_idx == -1:
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
                            # Use SMART extractor
                            description = _extract_middle_text_smart(str(desc_cell))
                        
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
                        if auth_idx is not None and auth_idx < len(row):
                            qty_cell = row[auth_idx]
                            if qty_cell:
                                match = re.search(r'(\d+)', str(qty_cell))
                                if match:
                                    qty = int(match.group(1))
                        
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


def generate_dd1750_from_pdf(bom_path: str, template_path: str, out_path: str, start_page: int = 0):
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"\nItems found: {len(items)}")
        
        if not items:
            return out_path, 0
        
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            packet = io.BytesIO()
            c = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
            
            first_row = Y_TABLE_TOP - 5.0
            
            for i, item in enumerate(page_items):
                y = first_row - (i * ROW_H)
                
                c.setFont("Helvetica", 8)
                c.drawCentredString((X_BOX_L + X_BOX_R) / 2, y - 7, str(item.line_no))
                
                c.setFont("Helvetica", 7)
                c.drawString(X_CONTENT_L + PAD_X, y - 7, item.description[:50])
                
                if item.nsn:
                    c.setFont("Helvetica", 6)
                    c.drawString(X_CONTENT_L + PAD_X, y - 12, f"NSN: {item.nsn}")
                
                c.setFont("Helvetica", 8)
                c.drawCentredString((X_UOI_L + X_UOI_R) / 2, y - 7, "EA")
                c.drawCentredString((X_INIT_L + X_INIT_R) / 2, y - 7, str(item.qty))
                c.drawCentredString((X_SPARES_L + X_SPARES_R) / 2, y - 7, "0")
                c.drawCentredString((X_TOTAL_L + X_TOTAL_R) / 2, y - 7, str(item.qty))
            
            c.save()
            packet.seek(0)
            
            overlay = PdfReader(packet)
            page = PdfReader(template_path).pages[page_num]
            page.merge_page(overlay.pages[0])
            writer.add_page(page)
    
    with open(out_path, 'wb') as f:
        writer.write(f)
    
    return out_path, len(items)
