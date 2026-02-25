import os
import logging
from io import BytesIO
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from app.schemas.individual_menu import IndividualSignRequest

logger = logging.getLogger(__name__)

LOGO_PATH = os.path.join(os.path.dirname(__file__), "..", "static", "logo_fdce.png")

class ItemData:
    def __init__(self, name, description, diet_options):
        self.name = name
        self.description = description
        self.diet_options = diet_options


def format_cell(cell, item):
    """Formats a table cell as a label."""
    # Set cell dimensions (approx 8cm x 5cm)
    # Note: python-docx doesn't always strictly enforce cell width if table is auto-layout
    # but we will try.
    
    # Clear cell
    for paragraph in cell.paragraphs:
        p = paragraph._element
        p.getparent().remove(p)
        
    # Vertical center
    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    
    # Name (Destacado)
    name_para = cell.add_paragraph()
    name_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = name_para.add_run(item.name.upper())
    run.font.name = "Arial"
    run.font.size = Pt(22)
    run.bold = True
    run.font.color.rgb = RGBColor(0x5a, 0x2d, 0x5a)
    
    # Description
    desc_para = cell.add_paragraph()
    desc_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = desc_para.add_run(item.description)
    run.font.name = "Arial"
    run.font.size = Pt(14)
    run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    
    # Diet Options
    if item.diet_options and item.diet_options.strip() and item.diet_options.strip() != '""' and item.diet_options.strip() != '"".""':
        diet_para = cell.add_paragraph()
        diet_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = diet_para.add_run(f"({item.diet_options.strip().upper()})")
        run.font.name = "Arial"
        run.font.size = Pt(12)
        run.bold = True
        run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

    # Logo (Bottom right)
    if os.path.exists(LOGO_PATH):
        logo_para = cell.add_paragraph()
        logo_para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run = logo_para.add_run()
        run.add_picture(LOGO_PATH, width=Cm(1.2))

def apply_borders_to_cell(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    
    for border_name in ['top', 'left', 'bottom', 'right']:
        tag = OxmlElement(f'w:{border_name}')
        tag.set(qn('w:val'), 'single')
        tag.set(qn('w:sz'), '12')
        tag.set(qn('w:space'), '0')
        tag.set(qn('w:color'), '5a2d5a')
        tcBorders.append(tag)
        
    tcPr.append(tcBorders)

def create_grid_page(doc):
    """Creates a 3x3 layout table grid for the page."""
    # We use a 3x3 layout to add a spacer column in the middle
    # and a merged center cell on the second row
    table = doc.add_table(rows=3, cols=3)
    table.autofit = False
    
    # Set row heights: Top (6cm), Middle (6cm), Bottom (6cm)
    for row in table.rows:
        row.height = Cm(6)
        
    # Set column widths. Total width ~= 27.7cm for Landscape A4 minus margins
    table.columns[0].width = Cm(11.5)
    table.columns[1].width = Cm(4.7) # Spacer
    table.columns[2].width = Cm(11.5)
    
    # Merge middle row for the center element
    table.cell(1, 0).merge(table.cell(1, 2))
    
    return table


def generate_individual_signs_docx(request: IndividualSignRequest) -> BytesIO:
    doc = Document()
    
    # Set Page Orientation to Landscape
    section = doc.sections[0]
    section.orientation = WD_ORIENT.LANDSCAPE
    # Update height/width for landscape
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height
    
    # Margins (1cm)
    section.top_margin = Cm(1)
    section.bottom_margin = Cm(1)
    section.left_margin = Cm(1)
    section.right_margin = Cm(1)
    
    items = []
    
    def process_and_add_item(name, desc, diet):
        if name and name.strip():
            # Check if name contains commas and split it
            if "," in name:
                name_parts = [p.strip() for p in name.split(",") if p.strip()]
                # If desc has commas, split it too, to map 1:1 with names. 
                # Otherwise, treat it as empty or use the whole thing if it's only 1 part (which shouldn't happen if names > 1, but just in case)
                desc_parts = []
                if desc:
                    desc_parts = [p.strip() for p in desc.split(",")]
                
                diet_parts = []
                if diet:
                    # Sometimes diet is "(GF), (VG)" or just "GF, VG". We split by comma if present.
                    # But diet often applies to all, or is mapped 1:1. We assume 1:1 if it has commas.
                    diet_parts = [p.strip() for p in diet.split(",")]

                for i, part in enumerate(name_parts):
                    # Try to get the matching description, fallback to empty string
                    part_desc = desc_parts[i] if i < len(desc_parts) else ""
                    # If there's only one description but many names, it might be a global description. 
                    # But per user request, we want to split. If it's missing, leave it blank.
                    if not part_desc and len(desc_parts) == 1:
                         # Edge case: If there's exactly 1 description but 4 names, do we repeat or leave blank?
                         # The user specifically complained it repeated. So we leave it blank if no 1:1 match.
                         # Actually, wait. If `Dried Apricots, Dried Cramberri` is the desc, it splits to 2.
                         # If `APRICOTS, CRAMBERRI` is the name, it splits to 2. Perfect match.
                         part_desc = desc_parts[0] if len(desc_parts) == 1 else ""

                    part_diet = diet_parts[i] if i < len(diet_parts) else ""
                    if not part_diet and len(diet_parts) == 1:
                        part_diet = diet_parts[0]
                        
                    items.append(ItemData(
                        part,
                        part_desc, 
                        part_diet
                    ))
            else:
                items.append(ItemData(
                    name.strip(), 
                    desc.strip() if desc else "", 
                    diet.strip() if diet else ""
                ))

    for meal in request.meals:
        process_and_add_item(meal.menu_name, meal.menu_desc, meal.menu_diet)
            
        for i in range(1, 11):
            name = getattr(meal, f"menu_{i}_name")
            desc = getattr(meal, f"menu_{i}_desc")
            diet = getattr(meal, f"menu_{i}_diet")
            process_and_add_item(name, desc, diet)
    
    total_items = len(items)
    
    for i in range(0, total_items, 5):
        if i > 0:
            doc.add_page_break()
            
        table = create_grid_page(doc)
        page_items = items[i:i+5]
        
        # Pattern:
        # R0C0 (Item 1), R0C1 (Item 2)
        # R1 (Merged) (Item 3)
        # R2C0 (Item 4), R2C1 (Item 5)
        
        # Mapping indices to table cells
        # cell_map: (page_index) -> (row, col)
        cell_map = {
            0: (0, 0),
            1: (0, 2),
            2: (1, 0), # This is the merged center cell
            3: (2, 0),
            4: (2, 2)
        }
        
        for idx, item in enumerate(page_items):
            row, col = cell_map[idx]
            outer_cell = table.cell(row, col)
            outer_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            # Clear outer cell default paragraph to avoid weird margins
            for paragraph in outer_cell.paragraphs:
                p = paragraph._element
                p.getparent().remove(p)
                
            # Create a nested 1x1 table for the actual card
            inner_table = outer_cell.add_table(rows=1, cols=1)
            inner_table.autofit = False
            inner_table.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            # Fixed dimensions for the card
            inner_table.columns[0].width = Cm(10)
            inner_table.rows[0].height = Cm(5.5)
            
            inner_cell = inner_table.cell(0, 0)
            
            # Apply borders to the nested table cell
            apply_borders_to_cell(inner_cell)
            
            # Format text inside the card
            format_cell(inner_cell, item)
            
    # Save to memory stream
    docx_stream = BytesIO()
    doc.save(docx_stream)
    docx_stream.seek(0)
    
    return docx_stream
