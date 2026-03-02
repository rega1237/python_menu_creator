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

        # Replace in tables (Crucial for the user's current layout)
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

        # Find the insertion point (paragraph containing the marker)
        marker_para = None
        for p in doc.paragraphs:
            if "[DYNAMIC_CONTENT_START]" in p.text:
                marker_para = p
                break

        # Create the main 2-column layout table
        # We'll create it first, then move it in the XML
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        
        sidebar_cell = table.cell(0, 0)
        main_cell = table.cell(0, 1)
        
        # Explicitly set column widths
        table.columns[0].width = Cm(5.5)
        table.columns[1].width = Cm(12.5)

        self._add_sidebar_content(sidebar_cell, request)
        
        # --- Main Content (Right Side) ---
        main_cell.width = Cm(12.5)
        
        # ... (rest of the content building logic stays the same) ...
        # [I will keep the logic here for the example, but it's identical to before]
        # Skip down to where we move the table.

        # ... (Assuming the rest of the dynamic content building happened here) ...
        # (Actually, let's keep the full method logic to ensure it works)
        
        hp = main_cell.add_paragraph()
        hr = hp.add_run("PROPOSAL OF SERVICES")
        hr.bold = True
        hr.font.size = Pt(14)
        hr.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
        
        main_cell.add_paragraph(request.event.end_date_formatted)
        main_cell.add_paragraph()

        # Meals Loop
        if request.meals:
            # ... meals content ...
            for meal in request.meals:
                if meal.show_date_header:
                    dp = main_cell.add_paragraph()
                    dr = dp.add_run(meal.date_header)
                    dr.bold = True
                
                main_cell.add_paragraph() # Spacer
                cat_p = main_cell.add_paragraph()
                cat_time = f" ({meal.time_range})" if meal.time_range else ""
                cat_run = cat_p.add_run(f"{meal.category_name}{cat_time}")
                cat_run.bold = True
                cat_run.font.color.rgb = RGBColor(0x61, 0x2d, 0x4b)
                
                if meal.description:
                    main_cell.add_paragraph(meal.description)
                
                for sub in meal.subcategories:
                    # Skip empty subcategories (common in flat AppSheet structures)
                    if not sub.name.strip() and not sub.menu_list.strip():
                        continue

                    sub_p = main_cell.add_paragraph()
                    sub_r = sub_p.add_run(sub.name)
                    sub_r.bold = True
                    sub_r.underline = True
                    sub_r.font.color.rgb = RGBColor(0, 0, 0)
                    
                    if sub.description:
                        main_cell.add_paragraph(sub.description)
                    
                    if sub.menu_list:
                        # Extract items using a strictly defined delimiter "|ITEM|"
                        # Commas are prevalent in names, descriptions, and diet options,
                        # so splitting by comma natively is too unsafe.
                        # AppSheet will send multiple items with the default " , " joined list,
                        # but we must use SUBSTITUTE in AppSheet to replace " , " with " |ITEM| "
                        raw_menus = sub.menu_list.split(" |ITEM| ")
                        menus = [m.strip() for m in raw_menus if m.strip()]
                        
                        for menu_item in menus:
                            # Parse format: "Name || Description || Diet Options"
                            parts = [p.strip() for p in menu_item.split("||")]
                            name = parts[0] if len(parts) > 0 else menu_item
                            desc = parts[1] if len(parts) > 1 else ""
                            diet = parts[2] if len(parts) > 2 else ""

                            menu_p = main_cell.add_paragraph()
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
                                desc_p = main_cell.add_paragraph()
                                desc_p.paragraph_format.left_indent = Cm(1.0)
                                desc_run = desc_p.add_run(desc)
                                desc_run.italic = True
                                desc_run.font.size = Pt(10)
                                desc_run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
                
                main_cell.add_paragraph()

        # Financials
        fin_title = main_cell.add_paragraph()
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
            p = main_cell.add_paragraph()
            p.add_run(f"{label}: ").bold = True
            p.add_run(val)
            
        total_p = main_cell.add_paragraph()
        total_r = total_p.add_run(f"Total Estimate: {request.financials.total_estimate}")
        total_r.bold = True
        total_r.font.size = Pt(13)

        # MOVE TABLE to marker position
        if marker_para:
            # Move the table before the marker paragraph
            marker_para._p.addnext(table._tbl)
            # Remove the marker paragraph
            p_element = marker_para._element
            p_element.getparent().remove(p_element)

        # Save to memory stream
        docx_stream = BytesIO()
        doc.save(docx_stream)
        docx_stream.seek(0)
        return docx_stream
