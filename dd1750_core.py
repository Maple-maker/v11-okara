"""DD1750 core - Complete working version."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Dict

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.lib.pagesizes import letter


PAGE_W, PAGE_H = letter

# === ADMIN FIELD POSITIONS (TOP SECTION) ===
# Adjust these Y values based on your template
ADMIN_COORDS = {
    'unit': {'x': 50, 'y': 735},           # UNIT box
    'requisition': {'x': 280, 'y': 735},   # REQUISITION NO. box
    'page': {'x': 500, 'y': 735},          # PAGE box
    'date': {'x': 50, 'y': 710},           # DATE box
    'order': {'x': 280, 'y': 710},         # ORDER NO. box
    'boxes': {'x': 480, 'y': 710},         # TOTAL NO. OF BOXES box
    'packed_by': {'x': 44, 'y': 115},      # PACKED BY (on every page)
    'received_by': {'x': 300, 'y': 115},   # RECEIVED BY (on every page)
}

# Table positions
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5
ROWS_PER_PAGE = 18
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


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
                            elif 'DESC' in text or 'NOMENCLATURE' in text:
                                desc_idx = i
                            elif 'MATERIAL' in text:
                                mat_idx = i
                            elif 'AUTH' in text or 'QTY' in text or 'QUANTITY' in text:
                                qty_idx = i
                    
                    if lv_idx is None or desc_idx is None:
                        continue
                    
                    for row in table[1:]:
                        if not any(cell for cell in row if cell):
                            continue
                        
                        lv_cell = row[lv_idx] if lv_idx < len(row) else None
                        if not lv_cell or str(lv_cell).strip().upper() != 'B':
                            continue
                        
                        # Get description
                        desc_cell = row[desc_idx] if desc_idx < len(row) else None
                        description = ""
                        if desc_cell:
                            lines = str(desc_cell).strip().split('\n')
                            description = lines[1].strip() if len(lines) >= 2 else lines[0].strip()
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
                        
                        # Get quantity - Try column first
                        qty = 1
                        
                        # Strategy 1: Qty column
                        if qty_idx is not None and qty_idx < len(row):
                            qty_cell = row[qty_idx]
                            if qty_cell:
                                qty_text = str(qty_cell).strip()
                                qty_match = re.search(r'(\d+)', qty_text)
                                if qty_match:
                                    qty = int(qty_match.group(1))
                        
                        # Strategy 2: Last word of description if it's a number
                        if qty == 1:
                            parts = description.split()
                            if parts and parts[-1].isdigit():
                                qty = int(parts[-1])
                                description = ' '.join(parts[:-1])
                        
                        items.append(BomItem(len(items) + 1, description[:100], nsn, qty))
    
    except Exception as e:
        print(f"ERROR: {e}")
    
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0
