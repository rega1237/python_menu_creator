from fastapi import FastAPI, Response, Query
import logging
import json

# Configure logging at the entry point
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

from app.schemas.menu import MenuRequest
from app.schemas.individual_menu import IndividualSignRequest
from app.services.general_sign_generator import generate_general_sign_docx
from app.services.individual_sign_generator import generate_individual_signs_docx
from app.services.google_drive_service import drive_service
from app.services.appsheet_service import appsheet_service
from app.services.estimate_docx_generator import EstimateDocxGenerator
from app.schemas.estimate_total import EstimateTotalRequest
import io

app = FastAPI(title="Menu Creator Service")

@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.post("/api/v1/menus/generate", response_class=Response)
async def generate_menu(
    request: MenuRequest, 
    upload_to_drive: bool = Query(False, description="Si es True, sube el archivo a Google Drive y devuelve JSON con el link.")
):
    docx_stream = generate_general_sign_docx(request)
    
    # Request event_name or fallback, cleaning potential extensions
    clean_event_name = request.event_name.split(".")[0]
    safe_event_name = clean_event_name.replace(" ", "_")
    filename = f"Sing_general_{safe_event_name}.docx"
    
    if upload_to_drive:
        result = drive_service.upload_file(docx_stream, filename)
        if result["success"]:
            # Callback to AppSheet (using download_link for direct download in AppSheet)
            appsheet_result = appsheet_service.update_event_sign_link(
                event_id=request.event_id,
                view_link=result["download_link"]
            )
            result["appsheet_update"] = appsheet_result
            
            import json
            return Response(
                content=json.dumps(result),
                media_type="application/json"
            )
        else:
            from fastapi import HTTPException
            raise HTTPException(status_code=500, detail=f"Error al subir a Drive: {result['error']}")
    
    # Default behavior: Return raw file
    docx_stream.seek(0)
    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }
    
    return Response(
        content=docx_stream.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers
    )

@app.post("/api/v1/menus/generate/individual")
async def generate_individual_signs(
    request: IndividualSignRequest,
    upload_to_drive: bool = Query(True)
):
    docx_stream = generate_individual_signs_docx(request)
    
    clean_event_name = request.event_name.split(".")[0]
    safe_event_name = clean_event_name.replace(" ", "_")
    filename = f"Individual_signs_{safe_event_name}.docx"
    
    if upload_to_drive:
        result = drive_service.upload_file(docx_stream, filename)
        if result["success"]:
            # Callback to AppSheet (using download_link and the INDIVIDUAL column)
            appsheet_result = appsheet_service.update_event_sign_link(
                event_id=request.event_id,
                view_link=result["download_link"],
                column_name="SINGS_INDIVIDUAL_WORD"
            )
            result["appsheet_update"] = appsheet_result
            
            return Response(
                content=json.dumps(result),
                media_type="application/json"
            )
            
    return Response(
        content=docx_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.post("/api/v1/menus/generate-estimate-total", response_class=Response)
async def generate_estimate_total(
    request: EstimateTotalRequest,
    upload_to_drive: bool = Query(True)
):
    """
    Generates a Word Estimate based on the provided JSON payload.
    """
    generator = EstimateDocxGenerator()
    docx_stream = generator.generate_docx(request)
    
    clean_event_name = request.event.name.split(".")[0] if request.event.name else "Unnamed"
    safe_event_name = clean_event_name.replace(" ", "_").replace("/", "-")
    filename = f"Estimate_{safe_event_name}.docx"

    if upload_to_drive:
        import json
        result = drive_service.upload_file(docx_stream, filename)
        if result["success"]:
            # Callback to AppSheet (Add new row to BDProposal History)
            appsheet_result = appsheet_service.add_proposal_history_row(
                event_id=request.event_id,
                doc_url=result["download_link"]
            )
            result["appsheet_update"] = appsheet_result
            
            return Response(
                content=json.dumps(result),
                media_type="application/json"
            )

    headers = {
        'Content-Disposition': f'attachment; filename="{filename}"'
    }
    
    return Response(
        content=docx_stream.getvalue(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers
    )
