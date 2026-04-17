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

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "..", "..", "estimate_template.docx")

class EstimateDocxGenerator:
    def __init__(self, template_path=TEMPLATE_PATH):
        self.template_path = template_path
        self.font_name = "Open Sans"
        self.primary_color = 0x612d4b  # Wine color from HTML
        self.text_color = 0x333333     # Main text color
        self.desc_color = 0x555555     # Description color

    def _set_run_font(self, run, size_pt=None, bold=None, italic=None, color_rgb=None, underline=None):
        """Helper to consistently set font properties in a run."""
        rPr = run._element.get_or_add_rPr()
        
        # Set font name in both high-level and XML level
        run.font.name = self.font_name
        rFonts = rPr.get_or_add_rFonts()
        rFonts.set(qn('w:ascii'), self.font_name)
        rFonts.set(qn('w:hAnsi'), self.font_name)
        rFonts.set(qn('w:cs'), self.font_name)

        if size_pt is not None:
            # size_pt is already in Pt/EMU if passed from Pt()
            # Directly assign to avoid double conversion bug
            run.font.size = size_pt
        if bold is not None:
            run.bold = bold
        if italic is not None:
            run.italic = italic
        if underline is not None:
            run.underline = underline
        if color_rgb is not None:
            run.font.color.rgb = RGBColor((color_rgb >> 16) & 0xff, (color_rgb >> 8) & 0xff, color_rgb & 0xff)

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
                                self._set_run_font(run) # Ensure replaced text matches font

        process_paragraphs(doc.paragraphs)

        def process_xml_element(element):
            for t_element in element.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"):
                if t_element.text:
                    for key, value in replacements.items():
                        if key in t_element.text:
                            t_element.text = t_element.text.replace(key, str(value or ""))
                            # For XML elements, we might need a different approach if they are broken
                            # but for now let's ensure the parent R has the font if possible.
                            parent_r = t_element.getparent()
                            if parent_r is not None:
                                rPr = parent_r.get_or_add_rPr()
                                rFonts = rPr.get_or_add_rFonts()
                                rFonts.set(qn('w:ascii'), self.font_name)
                                rFonts.set(qn('w:hAnsi'), self.font_name)

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

        def add_p(text="", alignment=None, space_after=Pt(6), space_before=Pt(0), bold=False, italic=False, size=Pt(11), color=0x333333, underline=False):
            if marker_para:
                p = marker_para.insert_paragraph_before(text)
            else:
                p = doc.add_paragraph(text)
            
            if alignment: p.alignment = alignment
            p.paragraph_format.space_after = space_after
            p.paragraph_format.space_before = space_before
            
            if p.runs:
                # Use current run if it exists, otherwise add one
                run = p.runs[0]
            else:
                run = p.add_run(text)
            
            # The add_run(text) might result in double text if p.runs already had text
            # but usually add_p(text) with marker_para doesn't have runs yet.
            if not p.runs and text:
                run = p.add_run(text)

            self._set_run_font(run, size_pt=size, bold=bold, italic=italic, color_rgb=color, underline=underline)
            
            return p

        def add_hr():
            p = add_p(space_after=Pt(2), space_before=Pt(12))
            p_pr = p._element.get_or_add_pPr()
            # Check if pBdr already exists to prevent duplication
            p_bdr = p_pr.find(qn('w:pBdr'))
            if p_bdr is None:
                p_bdr = OxmlElement('w:pBdr')
                p_pr.insert(0, p_bdr)
            
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '6')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), '000000')
            p_bdr.append(bottom)

        # --- MENU SECTION ---
        add_p("MENUS", bold=True, size=Pt(14), color=self.primary_color)
        add_p(request.event.date_formatted, size=Pt(11))

        if request.event.dietary_restrictions:
            add_p()
            add_p("Dietary Restrictions", bold=True, size=Pt(12), color=self.primary_color)
            add_p(request.event.dietary_restrictions)

        for meal in request.meals:
            add_hr()
            if meal.show_date_header:
                add_p(meal.date_header, bold=True, space_before=Pt(6))
            
            cat_text = meal.category_name.upper()
            if meal.time_range:
                cat_text += f": {meal.time_range}"
            
            add_p(cat_text, bold=True, size=Pt(13.5), color=self.primary_color, space_before=Pt(8))
            
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

                    if item.description:
                        desc_p = add_p(item.description, size=Pt(9), italic=True, color=self.desc_color, space_after=Pt(4))
                        desc_p.paragraph_format.left_indent = Cm(1.2)

        # --- FINANCIAL SECTION ---
        # Force a page break before financials if needed, or just a big spacer
        add_p(space_before=Pt(30))
        add_p("PROPOSAL OF SERVICES", bold=True, size=Pt(13.5), color=self.primary_color)
        add_p(request.event.end_date_formatted) # HTML uses End Event here
        add_p()

        # 1. Food Service
        add_p("Food Service", bold=True, size=Pt(12), color=self.primary_color)
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
            add_p("Labor Service Fees", bold=True, size=Pt(12), color=self.primary_color, space_before=Pt(15))
            
            # De-duplicate labor services based on content (preserve order)
            seen_labor = set()
            unique_labor = []
            for labor in request.labor_services:
                key = (labor.date_header, labor.hours, labor.name, labor.total)
                if key not in seen_labor:
                    seen_labor.add(key)
                    unique_labor.append(labor)

            for labor in unique_labor:
                if labor.show_date_header:
                    add_p(labor.date_header, bold=True, space_before=Pt(6))
                if labor.show_hours_header:
                    add_p(f"Staff suggested based on {labor.hours} hours of labor", size=Pt(10), italic=True)
                
                p = add_p(space_after=Pt(4))
                p.add_run(f"{labor.name}\t{labor.total}").bold = True

        # 3. Extras
        if request.extras_events:
            add_p("Extras Services", bold=True, size=Pt(12), color=self.primary_color, space_before=Pt(15))
            
            # De-duplicate extras events (preserve order)
            seen_extras = set()
            unique_extras = []
            for extra in request.extras_events:
                key = (extra.date_header, extra.is_rental, extra.is_sales, extra.name, extra.name_rental, extra.name_sales, extra.total, extra.provide_by_client)
                if key not in seen_extras:
                    seen_extras.add(key)
                    unique_extras.append(extra)

            for extra in unique_extras:
                if extra.show_date_header:
                    add_p(extra.date_header, bold=True)
                
                if extra.is_rental:
                    add_p("Rentals", bold=True, space_before=Pt(6))
                
                if extra.is_sales:
                    add_p("Sales", bold=True, space_before=Pt(6))

                p = add_p(space_after=Pt(2))
                # Determine name based on flags
                display_name = extra.name
                if not extra.provide_by_client:
                    if extra.is_rental: display_name = extra.name_rental
                    elif extra.is_sales: display_name = extra.name_sales

                if extra.provide_by_client:
                    txt = f"{display_name}\tProvide by the client"
                else:
                    txt = f"{display_name}\t{extra.total}"
                p.add_run(txt).bold = True

        # 4. Final Summary
        add_p(space_before=Pt(20))
        fin = request.financials
        
        def is_zero(val):
            if not val: return True
            clean_val = str(val).replace("$", "").replace(",", "").strip()
            try:
                return float(clean_val) == 0
            except ValueError:
                return False

        summary_items = [
            ("Food", fin.total_food_service, True),
            ("Labor Cost", fin.total_labor_cost, True),
            ("Extras Services", fin.total_extras_events, True),
            (f"{fin.tax_rate} {fin.tax_name}", fin.total_tax, True),
        ]
        
        # Add Extras Services (Sales) if not zero
        if not is_zero(fin.total_extras_sales):
            summary_items.append(("Extras Services", fin.total_extras_sales, True))
            
        summary_items.append((f"{fin.service_charge_rate} Service Charge", fin.total_service_charge, True))

        # Conditional items
        if fin.discount and not is_zero(fin.discount):
            val = fin.discount if str(fin.discount).startswith("-") else f"-{fin.discount}"
            summary_items.append(("Discount", val, False))
        if fin.donation and not is_zero(fin.donation):
            val = fin.donation if str(fin.donation).startswith("-") else f"-{fin.donation}"
            summary_items.append(("Donation", val, False))
        if fin.total_credit_card and not is_zero(fin.total_credit_card):
            summary_items.append(("Credit Card Fee", fin.total_credit_card, False))
        if fin.gratuity and not is_zero(fin.gratuity):
            summary_items.append(("Gratuity", fin.gratuity, False))

        for label, val, show_always in summary_items:
            p = add_p(space_after=Pt(2))
            
            # Setup tab stops for right alignment
            # Word default margin is usually around 16-17cm for Letter size
            tab_stops = p.paragraph_format.tab_stops
            tab_stops.add_tab_stop(Cm(16.5), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
            
            run = p.add_run(label)
            p.add_run("\t")
            
            if not is_zero(val):
                p.add_run(str(val))

        # Total Line
        total_p = add_p(space_after=Pt(2), space_before=Pt(8))
        tab_stops = total_p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Cm(16.5), alignment=WD_ALIGN_PARAGRAPH.RIGHT)
        
        r_total = total_p.add_run(f"Total Estimate\t{fin.total_estimate}")
        self._set_run_font(r_total, bold=True, size_pt=Pt(13), color_rgb=self.primary_color)
        
        p_pr = total_p._element.get_or_add_pPr()
        p_bdr = p_pr.find(qn('w:pBdr'))
        if p_bdr is None:
            p_bdr = OxmlElement('w:pBdr')
            p_pr.insert(0, p_bdr)
        
        top = OxmlElement('w:top')
        top.set(qn('w:val'), 'single')
        top.set(qn('w:sz'), '12')
        top.set(qn('w:color'), '612D4B')
        p_bdr.append(top)

        if marker_para:
            p_element = marker_para._element
            p_element.getparent().remove(p_element)

        docx_stream = BytesIO()
        doc.save(docx_stream)
        docx_stream.seek(0)
        return docx_stream
