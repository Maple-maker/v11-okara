"""DD1750 core - Extract middle text aligned with 'B' in LV column."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Tuple

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


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


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Extract items focusing on middle text aligned with 'B'."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page_num, page in enumerate(pdf.pages[start_page:], start=start_page):
                print(f"\n=== Processing page {page_num} ===")
                
                # Get all words with their positions
                words = page.extract_words()
                
                if not words:
                    continue
                
                # Group words by their y-position (rounded to nearest 0.5 for tolerance)
                y_groups = {}
                for word in words:
                    y = round(word['top'] * 2) / 2  # Round to nearest 0.5
                    if y not in y_groups:
                        y_groups[y] = []
                    y_groups[y].append(word)
                
                # Sort y positions from top to bottom
                sorted_y = sorted(y_groups.keys())
                
                # Find all "B" characters and their positions
                b_positions = []
                for y in sorted_y:
                    for word in y_groups[y]:
                        if word['text'].upper() == 'B':
                            b_positions.append({
                                'y': y,
                                'x': word['x0'],
                                'width': word['x1'] - word['x0'],
                                'text': word['text']
                            })
                
                print(f"  Found {len(b_positions)} 'B' characters")
                
                # For each "B", find text on the same line (same y-position)
                for b_idx, b_pos in enumerate(b_positions):
                    b_y = b_pos['y']
                    b_x = b_pos['x']
                    b_width = b_pos['width']
                    
                    # Get all words on the same horizontal line
                    same_line_words = []
                    for word in y_groups[b_y]:
                        # Skip the "B" itself
                        if word['text'].upper() == 'B':
                            continue
                        # Skip material codes and NSNs
                        if re.match(r'^\d{9}$', word['text']):
                            continue
                        if re.match(r'^C_[A-Z0-9]+$', word['text']):
                            continue
                        if word['text'] in ['WTY', 'ARC', 'CIIC', 'UI', 'SCMC', 'EA', 'AY', '9K', '9G']:
                            continue
                        
                        same_line_words.append(word)
                    
                    # Sort words by x-position (left to right)
                    same_line_words.sort(key=lambda w: w['x0'])
                    
                    # Join words to form description
                    description_parts = []
                    for word in same_line_words:
                        # Only include words that are to the right of "B"
                        if word['x0'] > b_x + b_width:
                            description_parts.append(word['text'])
                    
                    if description_parts:
                        description = ' '.join(description_parts)
                        
                        # Clean up the description
                        # Remove any parenthetical text
                        if '(' in description:
                            description = description.split('(')[0].strip()
                        
                        # Remove trailing codes
                        description = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC|EA|AY|9K|9G)\b$', '', description, flags=re.IGNORECASE)
                        description = re.sub(r'\s+', ' ', description).strip()
                        
                        # Look for NSN near this "B" position
                        nsn = ""
                        # Check words slightly above and below the "B"
                        for y in sorted_y:
                            if abs(y - b_y) <= 3.0:  # Check within 3 units vertically
                                for word in y_groups[y]:
                                    if re.match(r'^\d{9}$', word['text']):
                                        nsn = word['text']
                                        break
                            if nsn:
                                break
                        
                        # Look for quantity (usually a number at the end of the line)
                        qty = 1
                        for y in sorted_y:
                            if abs(y - b_y) <= 1.5:
                                for word in y_groups[y]:
                                    if word['text'].isdigit() and 1 <= int(word['text']) <= 100:
                                        qty = int(word['text'])
                                        break
                            if qty != 1:
                                break
                        
                        if description and len(description) > 2:
                            items.append(BomItem(
                                line_no=len(items) + 1,
                                description=description[:100],
                                nsn=nsn,
                                qty=qty
                            ))
                            print(f"  Item {len(items)}: '{description[:40]}...' | NSN: {nsn} | Qty: {qty}")
                
                # If position-based extraction didn't find enough items, fall back to table extraction
                if len(items) < 5:
                    print("  Falling back to table extraction")
                    tables = page.extract_tables()
                    
                    if tables:
                        for table in tables:
                            if len(table) < 2:
                                continue
                            
                            # Find column indices
                            header = table[0]
                            col_indices = {}
                            
                            for idx, cell in enumerate(header):
                                if cell:
                                    cell_text = str(cell).strip().upper()
                                    if 'LV' in cell_text or 'LEVEL' in cell_text:
                                        col_indices['lv'] = idx
                                    elif 'DESCRIPTION' in cell_text:
                                        col_indices['desc'] = idx
                                    elif 'MATERIAL' in cell_text:
                                        col_indices['material'] = idx
                                    elif 'QTY' in cell_text or 'QUANTITY' in cell_text:
                                        col_indices['qty'] = idx
                            
                            # Process rows
                            for row in table[1:]:
                                if 'lv' in col_indices and len(row) > col_indices['lv']:
                                    lv_cell = row[col_indices['lv']]
                                    if lv_cell and str(lv_cell).strip().upper() == 'B':
                                        # Get description
                                        description = ""
                                        if 'desc' in col_indices and len(row) > col_indices['desc']:
                                            desc_cell = row[col_indices['desc']]
                                            if desc_cell:
                                                # Split by newlines
                                                lines = str(desc_cell).strip().split('\n')
                                                if len(lines) >= 2:
                                                    # Take the second line (middle text)
                                                    description = lines[1].strip()
                                                else:
                                                    # Fallback: take everything
                                                    description = str(desc_cell).strip()
                                                
                                                # Clean up
                                                if ':' in description:
                                                    description = description.split(':')[0].strip()
                                                if '(' in description:
                                                    description = description.split('(')[0].strip()
                                                description = re.sub(r'\s+', ' ', description).strip()
                                        
                                        # Get NSN
                                        nsn = ""
                                        if 'material' in col_indices and len(row) > col_indices['material']:
                                            material_cell = row[col_indices['material']]
                                            if material_cell:
                                                material_text = str(material_cell).strip()
                                                nsn_match = re.search(r'\b(\d{9})\b', material_text)
                                                if nsn_match:
                                                    nsn = nsn_match.group(1)
                                        
                                        # Get quantity
                                        qty = 1
                                        if 'qty' in col_indices and len(row) > col_indices['qty']:
                                            qty_cell = row[col_indices['qty']]
                                            if qty_cell:
                                                try:
                                                    qty = int(str(qty_cell).strip())
                                                except:
                                                    qty = 1
                                        
                                        if description:
                                            items.append(BomItem(
                                                line_no=len(items) + 1,
                                                description=description[:100],
                                                nsn=nsn,
                                                qty=qty
                                            ))
                                            print(f"  Item {len(items)} (table): '{description[:40]}...'")
    
    except Exception as e:
        print(f"ERROR in extraction: {e}")
        return []
    
    print(f"\n=== Total items extracted: {len(items)} ===")
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 with proper multi-page support."""
    try:
        items = extract_items_from_pdf(bom_path, start_page)
        
        print(f"\n=== FINAL ITEM LIST ({len(items)} items) ===")
        for i, item in enumerate(items, 1):
            print(f"{i:3d}. '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
        print("=== END LIST ===\n")
        
        if not items:
            # Return empty template
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
            return output_path, 0
        
        # Calculate how many pages we need
        total_pages = math.ceil(len(items) / ROWS_PER_PAGE)
        print(f"Creating {total_pages} pages for {len(items)} items")
        
        writer = PdfWriter()
        
        for page_num in range(total_pages):
            print(f"  Creating page {page_num + 1}")
            
            # Get items for this page
            start_idx = page_num * ROWS_PER_PAGE
            end_idx = min((page_num + 1) * ROWS_PER_PAGE, len(items))
            page_items = items[start_idx:end_idx]
            
            # Create overlay for this page
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
            
            first_row_top = Y_TABLE_TOP_LINE - 5.0
            max_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
            
            for i, item in enumerate(page_items):
                y = first_row_top - (i * ROW_H)
                y_desc = y - 7.0
                y_nsn = y - 12.2
                
                # Box number
                can.setFont("Helvetica", 8)
                can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
                
                # Description
                can.setFont("Helvetica", 7)
                desc = item.description
                if len(desc) > 50:
                    desc = desc[:47] + "..."
                can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
                
                # NSN
                if item.nsn:
                    can.setFont("Helvetica", 6)
                    nsn_text = f"NSN: {item.nsn}"
                    can.drawString(X_CONTENT_L + PAD_X, y_nsn, nsn_text)
                
                # Quantities
                can.setFont("Helvetica", 8)
                can.drawCentredString((X_UOI_L + X_UOI_R)/2, y_desc, "EA")
                can.drawCentredString((X_INIT_L + X_INIT_R)/2, y_desc, str(item.qty))
                can.drawCentredString((X_SPARES_L + X_SPARES_R)/2, y_desc, "0")
                can.drawCentredString((X_TOTAL_L + X_TOTAL_R)/2, y_desc, str(item.qty))
            
            can.save()
            packet.seek(0)
            
            # Merge with template
            overlay = PdfReader(packet)
            template_page = PdfReader(template_path).pages[0]
            template_page.merge_page(overlay.pages[0])
            writer.add_page(template_page)
        
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        print(f"Successfully created {output_path} with {len(items)} items on {total_pages} pages")
        return output_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR: {e}")
        # Return empty template on any error
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
