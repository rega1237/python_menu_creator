import os
import logging
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
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

        # Replace in Headers (Important for the new repeated sidebar layout)
        for section in doc.sections:
            if section.header:
                process_paragraphs(section.header.paragraphs)
                for table in section.header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            process_paragraphs(cell.paragraphs)
            
            if section.first_page_header:
                process_paragraphs(section.first_page_header.paragraphs)
                for table in section.first_page_header.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            process_paragraphs(cell.paragraphs)

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

        def add_p(text=""):
            if marker_para:
                return marker_para.insert_paragraph_before(text)
            return doc.add_paragraph(text)

        # --- Main Content ---
        hp = add_p()
        hr = hp.add_run("PROPOSAL OF SERVICES")
        hr.bold = True
        hr.font.size = Pt(14)
        hr.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
        
        add_p(request.event.end_date_formatted)
        add_p()

        # Meals Loop
        if request.meals:
            # ... meals content ...
            for meal in request.meals:
                if meal.show_date_header:
                    dp = add_p()
                    dr = dp.add_run(meal.date_header)
                    dr.bold = True
                
                add_p() # Spacer
                cat_p = add_p()
                cat_time = f" ({meal.time_range})" if meal.time_range else ""
                cat_run = cat_p.add_run(f"{meal.category_name}{cat_time}")
                cat_run.bold = True
                cat_run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
                
                if meal.description:
                    add_p(meal.description)
                
                for sub in meal.subcategories:
                    # Skip empty subcategories (common in flat AppSheet structures)
                    if not sub.name.strip() and not sub.items:
                        continue

                    sub_p = add_p()
                    sub_r = sub_p.add_run(sub.name)
                    sub_r.bold = True
                    sub_r.underline = True
                    sub_r.font.color.rgb = RGBColor(0, 0, 0)
                    
                    if sub.description:
                        add_p(sub.description)
                    
                    if sub.items:
                        for menu_item in sub.items:
                            name = menu_item.name.strip()
                            desc = menu_item.description.strip()
                            diet = menu_item.diet_options.strip()

                            if not name and not desc:
                                continue

                            menu_p = add_p()
                            menu_p.paragraph_format.left_indent = Cm(0.5)
                            
                            # Add bullet and name (Purple, Bold)
                            menu_run = menu_p.add_run(f"• {name}")
                            menu_run.bold = True
                            menu_run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
                            
                            # Add diet details (Purple, Regular, Parentheses)
                            if diet:
                                diet_run = menu_p.add_run(f" ({diet})")
                                diet_run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
                            
                            # Add description (Gray, Italic, Smaller, Indented)
                            if desc:
                                desc_p = add_p()
                                desc_p.paragraph_format.left_indent = Cm(1.0)
                                desc_run = desc_p.add_run(desc)
                                desc_run.italic = True
                                desc_run.font.size = Pt(10)
                                desc_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
                
                add_p()

        # Financials
        fin_title = add_p()
        fin_run = fin_title.add_run("Estimation")
        fin_run.bold = True
        fin_run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
        
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
        total_r = total_p.add_run(f"Total Estimate: {request.financials.total_estimate}")
        total_r.bold = True
        total_r.font.size = Pt(13)

        # Remove the marker paragraph
        if marker_para:
            p_element = marker_para._element
            p_element.getparent().remove(p_element)

        # Save to memory stream
        docx_stream = BytesIO()
        doc.save(docx_stream)
        docx_stream.seek(0)
        return docx_stream
