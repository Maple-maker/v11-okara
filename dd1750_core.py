"""DD1750 core: parse BOM PDFs and render DD Form 1750 overlays.

SIMPLIFIED FIXED VERSION:
- 18 items per page (not 40)
- Only extract green-circled primary item names
- Ignore blue headers (COEI-, BII-)
- Ignore brown parenthetical text
- Simple, robust parsing
"""

from __future__ import annotations

import io
import math
import os
import re
from dataclasses import dataclass
from typing import List, Tuple

import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics


# --- Constants derived from the supplied blank template (letter: 612x792)
PAGE_W, PAGE_H = 612.0, 792.0

# Column x-bounds (points)
X_BOX_L, X_BOX_R = 44.0, 88.0
X_CONTENT_L, X_CONTENT_R = 88.0, 365.0
X_UOI_L, X_UOI_R = 365.0, 408.5
X_INIT_L, X_INIT_R = 408.5, 453.5
X_SPARES_L, X_SPARES_R = 453.5, 514.5
X_TOTAL_L, X_TOTAL_R = 514.5, 566.0

# Table y-bounds (points)
Y_TABLE_TOP_LINE = 616.0
Y_TABLE_BOTTOM_LINE = 89.5

# FIXED: Standard DD1750 has 18 lines, not 40
ROWS_PER_PAGE = 18

# Compute row height from table bounds
ROW_H = (Y_TABLE_TOP_LINE - Y_TABLE_BOTTOM_LINE) / ROWS_PER_PAGE

# Text padding inside cells
PAD_X = 3.0


@dataclass
class BomItem:
    line_no: int
    description: str
    nsn: str
    qty: int


def _extract_primary_description(text: str) -> str:
    """Extract only the green-circled primary item name."""
    if not text:
        return ""
    
    # Skip blue headers
    if text.startswith("COEI-") or text.startswith("BII-"):
        return ""
    
    # Remove any WTY, ARC, CIIC, UI, SCMC codes
    text = re.sub(r'\b(WTY|ARC|CIIC|UI|SCMC)\b.*?\d*', '', text, flags=re.IGNORECASE)
    
    # Remove material IDs like C_75Q65 ~ 1354640W
    text = re.sub(r'C_[A-Z0-9]+\s*~\s*[A-Z0-9]+', '', text)
    
    # Extract only up to first colon or parenthesis (green-circled text)
    # Stop at colon (for items like "BAG, TEXTILE: PAMPHLET")
    if ':' in text:
        text = text.split(':')[0].strip()
    
    # Stop at opening parenthesis (for brown parenthetical text)
    if '(' in text:
        text = text.split('(')[0].strip()
    
    # Remove any trailing commas, colons, or dashes
    text = re.sub(r'[,\-:]+\s*$', '', text)
    
    # Clean up whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    
    # Truncate if too long for DD1750
    if len(text) > 120:
        text = text[:117] + "..."
    
    return text


def extract_items_from_pdf(pdf_path: str, start_page: int = 0) -> List[BomItem]:
    """Simple, robust extraction of BOM items."""
    
    items: List[BomItem] = []
    
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = pdf.pages[start_page:]
            
            for page in pages:
                # Get all text
                text = page.extract_text() or ""
                if not text:
                    continue
                
                # Split into lines
                lines = [ln.strip() for ln in text.split('\n') if ln.strip()]
                
                i = 0
                while i < len(lines):
                    line = lines[i]
                    
                    # Skip blue headers
                    if line.startswith("COEI-") or line.startswith("BII-"):
                        i += 1
                        continue
                    
                    # Look for item patterns
                    # Check if line looks like an item description (contains comma, not just codes)
                    if (',' in line and 
                        len(line) > 5 and 
                        not re.match(r'^\d{9}$', line) and
                        not re.match(r'^[AB]$', line)):
                        
                        # Extract primary description (green-circled text)
                        description = _extract_primary_description(line)
                        
                        # Only process if we got a valid description
                        if description and len(description) > 3:
                            # Look for NSN in current or next lines
                            nsn = ""
                            qty = 1
                            
                            # Check current line for NSN
                            nsn_match = re.search(r'\b(\d{9})\b', line)
                            if nsn_match:
                                nsn = nsn_match.group(1)
                            else:
                                # Check next 3 lines for NSN
                                for j in range(1, 4):
                                    if i + j < len(lines):
                                        next_line = lines[i + j]
                                        nsn_match = re.search(r'\b(\d{9})\b', next_line)
                                        if nsn_match:
                                            nsn = nsn_match.group(1)
                                            break
                            
                            # Look for quantity (usually at end of line or in next lines)
                            qty_match = re.search(r'\b(\d+)\s*$', line)
                            if qty_match:
                                try:
                                    qty = int(qty_match.group(1))
                                except:
                                    qty = 1
                            
                            items.append(BomItem(
                                line_no=len(items) + 1,
                                description=description,
                                nsn=nsn,
                                qty=qty
                            ))
                    
                    i += 1
    
    except Exception as e:
        print(f"ERROR in extract_items_from_pdf: {e}")
        # Return empty list on error to prevent crash
        return []
    
    return items


def _wrap_to_width(text: str, font: str, size: float, max_w: float, max_lines: int) -> List[str]:
    """Simple word wrapping."""
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return [""]
    
    words = text.split(" ")
    lines: List[str] = []
    cur = ""
    
    def fits(s: str) -> bool:
        return pdfmetrics.stringWidth(s, font, size) <= max_w
    
    for w in words:
        if not cur:
            trial = w
        else:
            trial = cur + " " + w
        
        if fits(trial):
            cur = trial
            continue
        
        if cur:
            lines.append(cur)
            if len(lines) >= max_lines:
                return lines[:max_lines]
        
        if fits(w):
            cur = w
        else:
            # Word too long, break it
            for k in range(len(w), 0, -1):
                if fits(w[:k]):
                    lines.append(w[:k])
                    if len(lines) >= max_lines:
                        return lines[:max_lines]
                    cur = w[k:] if fits(w[k:]) else ""
                    break
    
    if cur and len(lines) < max_lines:
        lines.append(cur)
    
    return lines[:max_lines]


def _draw_center(c: canvas.Canvas, txt: str, x_l: float, x_r: float, y: float, font: str, size: float):
    c.setFont(font, size)
    x = (x_l + x_r) / 2.0
    c.drawCentredString(x, y, txt)


def _build_overlay_page(items: List[BomItem], page_num: int, total_pages: int) -> bytes:
    """Return a PDF bytes for a single overlay page with 18 rows max."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=(PAGE_W, PAGE_H))
    
    FONT_MAIN = "Helvetica"
    FONT_SMALL = "Helvetica"
    
    # Adjust baseline for 18 rows
    first_row_top = Y_TABLE_TOP_LINE - 5.0
    
    max_content_w = (X_CONTENT_R - X_CONTENT_L) - 2 * PAD_X
    
    for row_idx in range(ROWS_PER_PAGE):
        item_idx = row_idx
        y_row_top = first_row_top - row_idx * ROW_H
        y_desc = y_row_top - 7.0
        y_nsn = y_row_top - 12.2
        
        if item_idx >= len(items):
            continue
        
        it = items[item_idx]
        
        # Box number
        _draw_center(c, str(it.line_no), X_BOX_L, X_BOX_R, y_desc, FONT_MAIN, 8)
        
        # Description
        desc_lines = _wrap_to_width(it.description, FONT_MAIN, 7.0, max_content_w, max_lines=1)
        c.setFont(FONT_MAIN, 7.0)
        c.drawString(X_CONTENT_L + PAD_X, y_desc, desc_lines[0])
        
        # NSN if present
        if it.nsn:
            c.setFont(FONT_SMALL, 6.0)
            nsn_text = f"NSN: {it.nsn}"
            if pdfmetrics.stringWidth(nsn_text, FONT_SMALL, 6.0) <= max_content_w:
                c.drawString(X_CONTENT_L + PAD_X, y_nsn, nsn_text)
        
        # Quantities
        _draw_center(c, "EA", X_UOI_L, X_UOI_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_INIT_L, X_INIT_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, "0", X_SPARES_L, X_SPARES_R, y_desc, FONT_MAIN, 8)
        _draw_center(c, str(it.qty), X_TOTAL_L, X_TOTAL_R, y_desc, FONT_MAIN, 8)
    
    c.showPage()
    c.save()
    return buf.getvalue()


def generate_dd1750_from_pdf(
    bom_pdf_path: str,
    template_pdf_path: str,
    out_pdf_path: str,
    start_page: int = 0,
) -> Tuple[str, int]:
    """Generate DD1750 PDF - SIMPLIFIED to prevent crashes."""
    
    try:
        items = extract_items_from_pdf(bom_pdf_path, start_page=start_page)
        item_count = len(items)
        
        print(f"DEBUG: Successfully extracted {item_count} items")
        for i, item in enumerate(items[:5],
