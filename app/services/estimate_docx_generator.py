import os
import logging
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from app.schemas.estimate_total import EstimateTotalRequest

logger = logging.getLogger(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "estimate_temaplate.docx")

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

        def process_paragraphs(paragraphs):
            for paragraph in paragraphs:
                for key, value in replacements.items():
                    if key in paragraph.text:
                        for run in paragraph.runs:
                            if key in run.text:
                                run.text = run.text.replace(key, str(value or ""))

        process_paragraphs(doc.paragraphs)

        def process_xml_element(element):
            for t_element in element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                if t_element.text:
                    for key, value in replacements.items():
                        if key in t_element.text:
                            t_element.text = t_element.text.replace(key, str(value or ""))

        for section in doc.sections:
            if section.header:
                process_xml_element(section.header._element)
            if section.first_page_header:
                process_xml_element(section.first_page_header._element)

        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    process_paragraphs(cell.paragraphs)

    def generate_docx(self, request: EstimateTotalRequest) -> BytesIO:
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"Template not found at {self.template_path}")

        doc = Document(self.template_path)
        self._replace_placeholders(doc, request)

        marker_para = None
        for p in doc.paragraphs:
            if "[DYNAMIC_CONTENT_START]" in p.text:
                marker_para = p
                break

        if not marker_para:
            logger.warning("Marker [DYNAMIC_CONTENT_START] not found. Appending to end.")

        def add_p(text="", alignment=None, space_after=Pt(6), space_before=Pt(0), bold=False, italic=False, size=Pt(11), color=0x333333):
            if marker_para:
                p = marker_para.insert_paragraph_before(text)
            else:
                p = doc.add_paragraph(text)
            
            if alignment: p.alignment = alignment
            p.paragraph_format.space_after = space_after
            p.paragraph_format.space_before = space_before
            
            if text and p.runs:
                run = p.runs[0]
                run.font.name = "Open Sans"
                run.font.size = size
                run.bold = bold
                run.italic = italic
                if color is not None:
                    run.font.color.rgb = RGBColor((color >> 16) & 0xff, (color >> 8) & 0xff, color & 0xff)
            return p

        def add_hr():
            p = add_p(space_after=Pt(2), space_before=Pt(12))
            p_pr = p._element.get_or_add_pPr()
            p_bdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), '000000')
            p_bdr.append(bottom)
            p_pr.append(p_bdr)

        # --- MENU SECTION ---
        add_p("MENUS", bold=True, size=Pt(14), color=0x612d4b)
        add_p(request.event.date_formatted, size=Pt(11))

        if request.event.dietary_restrictions:
            add_p()
            add_p("Dietary Restrictions", bold=True, size=Pt(12), color=0x612d4b)
            add_p(request.event.dietary_restrictions)

        for meal in request.meals:
            add_hr()
            if meal.show_date_header:
                add_p(meal.date_header, bold=True, space_before=Pt(6))
            
            cat_text = meal.category_name.upper()
            if meal.time_range:
                cat_text += f": {meal.time_range}"
            
            add_p(cat_text, bold=True, size=Pt(14), color=0x612d4b, space_before=Pt(8))
            
            if meal.provide_by_client:
                p = add_p("◽ Provided by client", space_before=Pt(4))
                p.paragraph_format.left_indent = Cm(0.5)
                continue

            if meal.description:
                add_p(meal.description, italic=True)

            # --- MEAL LEVEL ---
            for i in range(1, 13):
                sub_name = getattr(meal, f"subcategory_{i}_name", "")
                sub_desc = getattr(meal, f"subcategory_{i}_description", "")
                sub_items = getattr(meal, f"subcategory_{i}_items", [])

                if not sub_name and not sub_items:
                    continue
                
                if sub_name:
                    sub_p = add_p(sub_name, bold=True, space_before=Pt(6))
                    sub_p.runs[0].underline = True

                if sub_desc:
                    add_p(sub_desc, size=Pt(10), italic=True)

                for item in sub_items:
                    # Construct the main item line
                    item_p = add_p(space_after=Pt(2))
                    item_p.paragraph_format.left_indent = Cm(0.8)
                    item_p.paragraph_format.first_line_indent = Cm(-0.4)
                    
                    r_bullet = item_p.add_run("◽ ")
                    r_bullet.bold = True
                    
                    r_name = item_p.add_run(item.name)
                    r_name.bold = True
                    r_name.underline = True
                    
                    if item.diet_options:
                        r_diet = item_p.add_run(f" ({item.diet_options})")
                        r_diet.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
                        r_diet.font.size = Pt(10)

                    if item.description:
                        desc_p = add_p(item.description, size=Pt(10), italic=True, color=0x555555, space_after=Pt(4))
                        desc_p.paragraph_format.left_indent = Cm(1.2)

        # --- FINANCIAL SECTION ---
        # Force a page break before financials if needed, or just a big spacer
        add_p(space_before=Pt(30))
        add_p("PROPOSAL OF SERVICES", bold=True, size=Pt(14), color=0x612d4b)
        add_p(request.event.end_date_formatted) # HTML uses End Event here
        add_p()

        # 1. Food Service
        add_p("Food Service", bold=True, size=Pt(12), color=0x612d4b)
        add_p(f"Based on {request.event.guests} Guests", size=Pt(10), italic=True)
        
        for meal in request.meals:
            if meal.show_date_header:
                add_p(meal.date_header, bold=True, space_before=Pt(6))
            
            p = add_p(space_after=Pt(2))
            r_label = p.add_run(meal.category_precio_guest)
            r_spacer = p.add_run("\t") # Tab to align right if possible, or just space
            if not meal.provide_by_client:
                r_val = p.add_run(meal.total_category_precio)
                r_val.bold = True
            else:
                p.add_run("Provided by client").italic = True

        # 2. Labor
        if request.labor_services:
            add_p("Labor Service Fees", bold=True, size=Pt(12), color=0x612d4b, space_before=Pt(15))
            for labor in request.labor_services:
                if labor.show_date_header:
                    add_p(labor.date_header, bold=True, space_before=Pt(6))
                if labor.show_hours_header:
                    add_p(f"Staff suggested based on {labor.hours} hours of labor", size=Pt(10), italic=True)
                
                p = add_p(space_after=Pt(4))
                p.add_run(f"{labor.name}\t{labor.total}").bold = True

        # 3. Extras
        if request.extras_events:
            add_p("Extras Services", bold=True, size=Pt(12), color=0x612d4b, space_before=Pt(15))
            for extra in request.extras_events:
                if extra.show_date_header:
                    add_p(extra.date_header, bold=True)
                
                p = add_p(space_after=Pt(2))
                txt = f"{extra.name}\t{extra.total}"
                if extra.provide_by_client: txt += " (Provided by client)"
                p.add_run(txt).bold = True

        # 4. Final Summary
        add_p(space_before=Pt(20))
        fin = request.financials
        summary_items = [
            ("Food", fin.total_food_service),
            ("Labor Cost", fin.total_labor_cost),
            ("Extras Services", fin.total_extras_events),
            (f"{fin.tax_rate} {fin.tax_name}", fin.total_tax),
            (f"{fin.service_charge_rate} Service Charge", fin.total_service_charge),
        ]
        if fin.discount and fin.discount != "0": summary_items.append(("Discount", f"-{fin.discount}"))
        if fin.donation and fin.donation != "0": summary_items.append(("Donation", f"-{fin.donation}"))
        if fin.total_credit_card and fin.total_credit_card != "0": summary_items.append(("Credit Card Fee", fin.total_credit_card))
        if fin.gratuity and fin.gratuity != "0": summary_items.append(("Gratuity", fin.gratuity))

        for label, val in summary_items:
            p = add_p(space_after=Pt(2))
            p.add_run(f"{label}\t{val}")

        # Total Line
        total_p = add_p(f"Total Estimate\t{fin.total_estimate}", bold=True, size=Pt(13), color=0x612d4b, space_before=Pt(8))
        p_pr = total_p._element.get_or_add_pPr()
        p_bdr = OxmlElement('w:pBdr')
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '12')
        top.set(qn('w:color'), '612D4B')
        p_bdr.append(top)
        p_pr.append(p_bdr)

        if marker_para:
            p_element = marker_para._element
            p_element.getparent().remove(p_element)

        docx_stream = BytesIO()
        doc.save(docx_stream)
        docx_stream.seek(0)
        return docx_stream
