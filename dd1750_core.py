"""DD1750 core - Extract only top/larger text from description boxes."""

import io
import math
import re
from dataclasses import dataclass
from typing import List, Tuple, Dict, Any

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
    """Extract items using text positioning to get only top/larger text."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                # Get words with their positions
                words = page.extract_words()
                
                if not words:
                    continue
                
                # Group words by row based on y-position
                # Words with similar y-coordinates are on the same row
                rows = {}
                for word in words:
                    y = round(word['top'], 1)  # Round to 1 decimal for grouping
                    if y not in rows:
                        rows[y] = []
                    rows[y].append(word)
                
                # Sort rows by y-position (top to bottom)
                sorted_y_positions = sorted(rows.keys(), reverse=True)  # PDF y=0 is bottom
                
                # Process each potential row
                current_item = None
                pending_nsn = None
                
                for y_pos in sorted_y_positions:
                    row_words = rows[y_pos]
                    
                    # Sort words by x-position (left to right)
                    row_words.sort(key=lambda w: w['x0'])
                    
                    # Join words to form row text
                    row_text = ' '.join([w['text'] for w in row_words])
                    
                    # Check if this row has LV = 'B'
                    # Look for "B" as a separate word or at start of line
                    if re.search(r'\bB\b', row_text) and len(row_text.split()) >= 3:
                        # This is likely an item row with LV = 'B'
                        
                        # Save previous item if we have one
                        if current_item and current_item.description:
                            items.append(current_item)
                        
                        # Extract description - take text after "B" and any material code
                        parts = row_text.split()
                        desc_start = 2  # Skip material code and "B"
                        if len(parts) > desc_start:
                            # Get all text after the first 2 columns
                            description_parts = parts[desc_start:]
                            
                            # Filter: only take words that look like description (not codes)
                            filtered_parts = []
                            for part in description_parts:
                                # Skip if it's a code (all caps, short, no vowels, etc.)
                                if (len(part) <= 3 and part.isupper()) or \
                                   re.match(r'^[A-Z0-9_\-]+$', part) or \
                                   part in ['WTY', 'ARC', 'CIIC', 'UI', 'SCMC', 'EA', 'AY', '9K', '9G']:
                                    continue
                                filtered_parts.append(part)
                            
                            description = ' '.join(filtered_parts)
                            
                            # Clean up
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Extract quantity (usually at end)
                            qty = 1
                            if description and description.split()[-1].isdigit():
                                qty = int(description.split()[-1])
                                description = ' '.join(description.split()[:-1])
                            
                            current_item = BomItem(
                                line_no=len(items) + 1,
                                description=description[:100],
                                nsn="",  # Will be filled from nearby text
                                qty=qty
                            )
                    
                    # Look for NSN (9-digit number) on any row
                    nsn_match = re.search(r'\b(\d{9})\b', row_text)
                    if nsn_match:
                        pending_nsn = nsn_match.group(1)
                        # If we have a current item, assign the NSN to it
                        if current_item and not current_item.nsn:
                            current_item.nsn = pending_nsn
                
                # Don't forget the last item
                if current_item and current_item.description:
                    items.append(current_item)
    
    except Exception as e:
        print(f"ERROR in extraction: {e}")
        return []
    
    return items


def extract_items_simple(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Simpler extraction for reliability."""
    items = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages[start_page:]:
                text = page.extract_text() or ""
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                
                i = 0
                while i < len(lines):
                    line = lines[i]
                    
                    # Look for lines with "B" (LV column) and description
                    # Pattern: material_code B description...
                    if re.match(r'^\S+\s+B\s+', line) or re.match(r'^B\s+', line):
                        parts = line.split()
                        
                        # Find where description starts (after material code and "B")
                        desc_start = 0
                        for idx, part in enumerate(parts):
                            if part == 'B' and idx > 0:  # Found "B" after material code
                                desc_start = idx + 1
                                break
                        
                        if desc_start > 0 and desc_start < len(parts):
                            # Extract description parts
                            desc_parts = []
                            for part in parts[desc_start:]:
                                # Skip if it looks like a code
                                if (len(part) <= 3 and part.isupper()) or \
                                   re.match(r'^[A-Z0-9_\-]+$', part) or \
                                   part in ['WTY', 'ARC', 'CIIC', 'UI', 'SCMC', 'EA', 'AY', '9K', '9G']:
                                    continue
                                desc_parts.append(part)
                            
                            description = ' '.join(desc_parts)
                            
                            # Clean: remove anything after colon or parenthesis
                            if ':' in description:
                                description = description.split(':')[0].strip()
                            if '(' in description:
                                description = description.split('(')[0].strip()
                            
                            description = re.sub(r'\s+', ' ', description).strip()
                            
                            # Extract quantity
                            qty = 1
                            if description and description.split()[-1].isdigit():
                                qty = int(description.split()[-1])
                                description = ' '.join(description.split()[:-1])
                            
                            # Look for NSN in current or next lines
                            nsn = ""
                            for j in range(max(0, i-1), min(len(lines), i+3)):
                                nsn_match = re.search(r'\b(\d{9})\b', lines[j])
                                if nsn_match:
                                    nsn = nsn_match.group(1)
                                    break
                            
                            if description:
                                items.append(BomItem(
                                    line_no=len(items) + 1,
                                    description=description[:100],
                                    nsn=nsn,
                                    qty=qty
                                ))
                    
                    i += 1
    
    except Exception as e:
        print(f"ERROR in simple extraction: {e}")
    
    return items


def generate_dd1750_from_pdf(bom_path, template_path, output_path, start_page=0):
    """Generate DD1750 - Try multiple extraction methods."""
    items = []
    
    # Try position-based extraction first
    items = extract_items_from_pdf(bom_path, start_page)
    
    # Fallback to simple extraction if first method fails
    if not items:
        items = extract_items_simple(bom_path, start_page)
    
    print(f"DEBUG: Found {len(items)} items")
    for i, item in enumerate(items[:10], 1):
        print(f"  Item {i}: '{item.description}' | NSN: {item.nsn} | Qty: {item.qty}")
    
    try:
        if not items:
            # Return empty template
            reader = PdfReader(template_path)
            writer = PdfWriter()
            writer.add_page(reader.pages[0])
            with open(output_path, 'wb') as f:
                writer.write(f)
            return output_path, 0
        
        # Create overlay
        packet = io.BytesIO()
        can = canvas.Canvas(packet, pagesize=(PAGE_W, PAGE_H))
        
        first_row_top = Y_TABLE_TOP_LINE - 5.0
        max_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
        
        for i in range(min(len(items), ROWS_PER_PAGE)):
            item = items[i]
            y = first_row_top - (i * ROW_H)
            y_desc = y - 7.0
            y_nsn = y - 12.2
            
            # Box number
            can.setFont("Helvetica", 8)
            can.drawCentredString((X_BOX_L + X_BOX_R)/2, y_desc, str(item.line_no))
            
            # Description (truncate if too long)
            can.setFont("Helvetica", 7)
            desc = item.description
            if len(desc) > 50:
                desc = desc[:47] + "..."
            can.drawString(X_CONTENT_L + PAD_X, y_desc, desc)
            
            # NSN
            if item.nsn:
                can.setFont("Helvetica", 6)
                nsn_text = f"NSN: {item.nsn}"
                if pdfmetrics.stringWidth(nsn_text, "Helvetica", 6) <= max_w:
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
        reader = PdfReader(template_path)
        writer = PdfWriter()
        
        overlay = PdfReader(packet)
        page = reader.pages[0]
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
        
        with open(output_path, 'wb') as f:
            writer.write(f)
        
        return output_path, len(items)
        
    except Exception as e:
        print(f"CRITICAL ERROR in generation: {e}")
        # Return empty template on any error
        reader = PdfReader(template_path)
        writer = PdfWriter()
        writer.add_page(reader.pages[0])
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path, 0
