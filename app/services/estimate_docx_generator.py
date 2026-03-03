import os
import logging
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_ALIGN_VERTICAL
from app.schemas.estimate_total import EstimateTotalRequest

logger = logging.getLogger(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "estimate_template.docx")
LOGO_PATH = os.path.join(os.path.dirname(__file__), "..", "static", "logo_fdce.png")
BANNER_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "banner.jpg")



class EstimateDocxGenerator:
    def __init__(self, template_path=TEMPLATE_PATH):
        self.template_path = template_path

    def _replace_placeholders(self, doc, request: EstimateTotalRequest):
        """Replaces {{TAG}} placeholders in the document paragraphs and tables."""
        replacements = {
            "{{EVENT_NAME}}": request.event.name,
            "{{CLIENT_NAME}}": request.client.name,
            "{{CLIENT_ADDRESS}}": request.client.address,
            "{{CLIENT_EMAIL}}": request.client.email,
            "{{REPRESENTATIVE_NAME}}": request.client_representative.name,
            "{{REPRESENTATIVE_EMAIL}}": request.client_representative.email,
            "{{REPRESENTATIVE_PHONE}}": request.client_representative.formatted_phone,
            "{{EVENT_ADDRESS}}": request.event.address,
            "{{EVENT_CODE}}": request.event.code,
            "{{EVENT_START}}": request.event.date_formatted,
            "{{EVENT_END}}": request.event.end_date_formatted,
            "{{EVENT_GUESTS}}": str(request.event.guests),
            "{{SERVICE_CHARGE_RATE}}": request.financials.service_charge_rate,
        }

        # Helper to process a list of paragraphs
        def process_paragraphs(paragraphs):
            for paragraph in paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        # Try replacement in runs first to preserve formatting and images
                        found_in_run = False
                        for run in paragraph.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, str(value or ""))
                                found_in_run = True
                        
                        # If not found in any single run (split across runs),
                        # we fall back to paragraph.text replacement.
                        if not found_in_run:
                            paragraph.text = paragraph.text.replace(key, str(value or ""))

        # Replace in main paragraphs
        process_paragraphs(doc.paragraphs)

        # Helper to process XML nodes (needed for Text Boxes / Shapes in Headers)
        def process_xml_element(element):
            # w:t is the WordprocessingML tag for text
            for t_element in element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                if t_element.text:
                    for key, value in replacements.items():
                        if key in t_element.text:
                            t_element.text = t_element.text.replace(key, str(value or ""))

        # Replace in Headers (Important for the new repeated sidebar layout)
        # Includes standard paragraphs, tables, AND floating shapes (Text Boxes)
        for section in doc.sections:
            if section.header:
                process_xml_element(section.header._element)
            
            if section.first_page_header:
                process_xml_element(section.first_page_header._element)

        # Replace in tables in the main body
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_paragraphs(cell.paragraphs)

    def _add_sidebar_content(self, cell, request: EstimateTotalRequest):
        """Fills the left column (sidebar) with event and client details."""
        cell.width = Cm(5.5)
        
        # Banner at the top
        if os.path.exists(BANNER_PATH):
            banner_para = cell.add_paragraph()
            banner_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = banner_para.add_run()
            run.add_picture(BANNER_PATH, width=Cm(5.0))
            cell.add_paragraph() # Spacer

        # Tagline
        p = cell.add_paragraph()
        run = p.add_run("We create everlasting moments")
        run.font.size = Pt(13)
        run.bold = True
        run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
        run.italic = True

        # Sections
        sections = [
            ("Client", [request.client.name, request.client.address, request.client.email]),
            ("Client Representative", [
                request.client_representative.name, 
                request.client_representative.email,
                f"Ph: {request.client_representative.formatted_phone}"
            ]),
            ("Event Details", [
                request.event.name,
                request.event.address,
                f"Event Code: {request.event.code}",
                f"Start: {request.event.date_formatted}",
                f"End: {request.event.end_date_formatted}",
                f"{request.event.guests} Guests"
            ])
        ]

        for title, lines in sections:
            cell.add_paragraph() # Spacer
            tp = cell.add_paragraph()
            tr = tp.add_run(title)
            tr.bold = True
            tr.font.size = Pt(12)
            tr.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
            
            for line in lines:
                if line:
                    lp = cell.add_paragraph(line)
                    lp.paragraph_format.space_after = Pt(2)
                    run = lp.runs[0]
                    run.font.size = Pt(10)

        # Logo at the bottom
        if os.path.exists(LOGO_PATH):
            cell.add_paragraph() # Spacer
            logo_para = cell.add_paragraph()
            logo_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = logo_para.add_run()
            run.add_picture(LOGO_PATH, width=Cm(3.5))

    def generate_docx(self, request: EstimateTotalRequest) -> BytesIO:
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")

        doc = Document(self.template_path)
        self._replace_placeholders(doc, request)

        # 1. Find the insertion point (paragraph containing the marker)
        marker_para = None
        for p in doc.paragraphs:
            if "[DYNAMIC_CONTENT_START]" in p.text:
                marker_para = p
                break

        if not marker_para:
            logger.warning("Marker [DYNAMIC_CONTENT_START] not found in template. Appending to end.")
            container = doc
        else:
            container = marker_para # We'll use insert_paragraph_before logic

        def format_full_date(date_str: str) -> str:
            if not date_str:
                return ""
            try:
                # Try to parse MM/DD/YY or MM/DD/YYYY
                parts = date_str.split("/")
                if len(parts) == 3:
                    m, d, y = map(int, parts)
                    if y < 100:
                        y += 2000
                    dt = datetime(y, m, d)
                else:
                    # Try YYYY-MM-DD
                    dt = datetime.strptime(date_str, "%Y-%m-%d")
                
                # Format: Tuesday, October, 27th 2026
                day = dt.day
                if 11 <= day <= 13:
                    suffix = "th"
                else:
                    suffix = {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
                
                return dt.strftime(f"%A, %B, {day}{suffix} %Y")
            except Exception as e:
                logger.warning(f"Could not parse date '{date_str}': {e}")
                return date_str

        def add_p(text="", alignment=None, line_spacing=1.4, space_after=Pt(6)):
            if marker_para:
                p = marker_para.insert_paragraph_before(text)
            else:
                p = doc.add_paragraph(text)
            
            if alignment:
                p.alignment = alignment
            
            p.paragraph_format.line_spacing = line_spacing
            p.paragraph_format.space_after = space_after
            return p

        def style_run(run, font_name="Open Sans", size=None, bold=False, italic=False, underline=False, color=0x333333):
            run.font.name = font_name
            if size:
                run.font.size = size
            run.bold = bold
            run.italic = italic
            run.underline = underline
            if color is not None:
                run.font.color.rgb = RGBColor((color >> 16) & 0xff, (color >> 8) & 0xff, color & 0xff)
            return run

        # --- Main Content ---
        # Replace "PROPOSAL OF SERVICES" with Event Name
        hp = add_p(alignment=WD_ALIGN_PARAGRAPH.CENTER)
        hr = hp.add_run(request.event.name.upper() if request.event.name else "PROPOSAL OF SERVICES")
        style_run(hr, size=Pt(14), bold=True, color=0x612d4b)
        
        # Add Event Date formatted
        date_formatted = format_full_date(request.event.date_formatted)
        if date_formatted:
            dp = add_p(date_formatted, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            style_run(dp.runs[0], size=Pt(11))
        add_p()

        # Meals Loop
        if request.meals:
            # ... meals content ...
            for meal in request.meals:
                if meal.show_date_header:
                    dp = add_p()
                    dr = dp.add_run(format_full_date(meal.date_header) if meal.date_header else "")
                    style_run(dr, bold=True)
                
                add_p() # Spacer
                cat_p = add_p()
                cat_time = f" ({meal.time_range})" if meal.time_range else ""
                cat_run = cat_p.add_run(f"{meal.category_name.upper()}{cat_time}")
                style_run(cat_run, size=Pt(14), bold=True, color=0x612d4b)
                
                if meal.description:
                    p = add_p(meal.description)
                    style_run(p.runs[0])
                
                for sub in meal.subcategories:
                    # Skip empty subcategories (common in flat AppSheet structures)
                    if not sub.name.strip() and not sub.items:
                        continue

                    sub_p = add_p()
                    sub_r = sub_p.add_run(sub.name)
                    style_run(sub_r, bold=True, underline=True, color=0x333333)
                    
                    if sub.description:
                        p = add_p(sub.description)
                        style_run(p.runs[0])
                    
                    if sub.items:
                        for menu_item in sub.items:
                            name = menu_item.name.strip()
                            desc = menu_item.description.strip()
                            diet = menu_item.diet_options.strip()

                            if not name and not desc:
                                continue

                            menu_p = add_p(space_after=Pt(2))
                            # Hanging indent: indent whole paragraph, outdent first line
                            menu_p.paragraph_format.left_indent = Cm(1.0) # Overall indent (where text wraps)
                            menu_p.paragraph_format.first_line_indent = Cm(-0.5) # Bullet starts 0.5cm LEFT of that
                            
                            # Add bullet (Not underlined)
                            bullet_run = menu_p.add_run("• ")
                            style_run(bullet_run, size=Pt(13), bold=True)
                            
                            # Add name (Standard body color, Underlined per HTML)
                            menu_run = menu_p.add_run(name)
                            style_run(menu_run, size=Pt(13), bold=True, underline=True)
                            
                            # Add diet details (Purple, Regular, Parentheses)
                            if diet:
                                diet_run = menu_p.add_run(f" ({diet})")
                                style_run(diet_run, size=Pt(11), color=0x612d4b)
                            
                            # Add description (Gray, Italic, Smaller, Indented)
                            if desc:
                                desc_p = add_p(space_after=Pt(8))
                                desc_p.paragraph_format.left_indent = Cm(1.0)
                                desc_run = desc_p.add_run(desc)
                                style_run(desc_run, size=Pt(10), italic=True, color=0x555555)
                
                add_p()

        # Financials
        fin_title = add_p()
        fin_run = fin_title.add_run("Estimation")
        style_run(fin_run, size=Pt(14), bold=True, color=0x612d4b)
        
        for label, val in [
            ("Food", request.financials.total_food_service),
            ("Labor Cost", request.financials.total_labor_cost),
            ("Extras Services", request.financials.total_extras_events),
            (f"{request.financials.tax_rate} {request.financials.tax_name}", request.financials.total_tax),
            (f"{request.financials.service_charge_rate} Service Charge", request.financials.total_service_charge)
        ]:
            p = add_p()
            p.add_run(f"{label}: ").bold = True
            p.add_run(val)
            
        total_p = add_p()
        # Add a top border manually via XML for the total
        from docx.oxml.ns import qn
        from docx.oxml import OxmlElement
        p_pr = total_p._element.get_or_add_pPr()
        p_bdr = OxmlElement('w:pBdr')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '12') # 1.5 pt
        top.set(qn('w:space'), '4')
        top.set(qn('w:color'), '612D4B')
        p_bdr.append(top)
        p_pr.append(p_bdr)

        total_r = total_p.add_run(f"Total Estimate: {request.financials.total_estimate}")
        style_run(total_r, size=Pt(13), bold=True, color=0x612d4b)

        # Remove the marker paragraph
        if marker_para:
            p_element = marker_para._element
            p_element.getparent().remove(p_element)

        # Save to memory stream
        docx_stream = BytesIO()
        doc.save(docx_stream)
        docx_stream.seek(0)
        return docx_stream
