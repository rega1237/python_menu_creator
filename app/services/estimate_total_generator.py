import os
from jinja2 import Environment, FileSystemLoader
from weasyprint import HTML
import logging
from app.schemas.estimate_total import EstimateTotalRequest

logger = logging.getLogger(__name__)

class EstimateTotalGenerator:
    def __init__(self, template_dir="app/templates", template_name="estimate_total.html"):
        self.template_dir = template_dir
        self.template_name = template_name
        self._env = Environment(loader=FileSystemLoader(self.template_dir))

    def generate_pdf(self, request_data: EstimateTotalRequest) -> bytes:
        """
        Generates the Estimate Total PDF from the provided JSON request data.
        """
        try:
            # Load template
            template = self._env.get_template(self.template_name)

            # Render HTML with data
            # Convert Pydantic model to dict for jinja consumption
            html_content = template.render(**request_data.model_dump())

            # Generate PDF
            pdf_bytes = HTML(string=html_content).write_pdf()
            
            return pdf_bytes
            
        except Exception as e:
            logger.error(f"Error generating Estimate Total PDF: {e}")
            raise e
