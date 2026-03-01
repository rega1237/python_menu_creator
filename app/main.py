from fastapi import FastAPI, Response, Query, Request
from fastapi.responses import JSONResponse
from fastapi.exceptions import RequestValidationError
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
from app.schemas.excel_menu import ExcelMenuRequest
from app.services.excel_generator import generate_individual_excel, generate_combined_excel
import io

app = FastAPI(title="Menu Creator Service")

@app.exception_handler(RequestValidationError)
async def validation_exception_handler(request: Request, exc: RequestValidationError):
    """
    Custom handler to log the raw body when validation fails.
    This helps identify if binary data (like a PDF) is being sent instead of JSON.
    """
    body = await request.body()
    print(f"--- VALIDATION ERROR ---")
    print(f"Method: {request.method} URL: {request.url}")
    print(f"Headers: {request.headers}")
    try:
        # Try to show body as text for debugging
        print(f"Body Preview: {body[:200].decode('utf-8', errors='replace')}")
    except:
        print(f"Body Preview (Binary): {body[:200]}")
    
    return JSONResponse(
        status_code=422,
        content={
            "detail": exc.errors(),
            "message": "The server expected JSON but received something else (maybe binary data or a PDF). Check your AppSheet Webhook 'Body' configuration."
        }
    )

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

@app.post("/api/v1/menus/generate-excel", response_class=Response)
async def generate_excel_endpoint(
    request: ExcelMenuRequest,
    upload_to_drive: bool = Query(True)
):
    """
    Generates two Excel documents (Individual and Combined) and uploads them to Drive.
    Updates AppSheet BDEvents with both links if upload_to_drive is True.
    """
    import json
    
    # 1. Generate Individual Excel
    individual_stream = generate_individual_excel(request)
    # 2. Generate Combined Excel
    combined_stream = generate_combined_excel(request)
    
    clean_event_name = request.event_name.split(".")[0] if request.event_name else "Unnamed"
    safe_event_name = clean_event_name.replace(" ", "_").replace("/", "-")
    
    individual_filename = f"Individual_Excel_{safe_event_name}.xlsx"
    combined_filename = f"Combined_Excel_{safe_event_name}.xlsx"
    
    if upload_to_drive:
        ind_result = drive_service.upload_file(individual_stream, individual_filename)
        comb_result = drive_service.upload_file(combined_stream, combined_filename)
        
        response_data = {
            "individual_excel": ind_result,
            "combined_excel": comb_result,
            "success": ind_result.get("success", False) and comb_result.get("success", False)
        }
        
        if response_data["success"]:
            # Perform AppSheet update in the BDEvents table
            appsheet_result_ind = appsheet_service.update_event_sign_link(
                event_id=request.event_id,
                view_link=ind_result["download_link"],
                column_name="excel_individual"
            )
            
            appsheet_result_comb = appsheet_service.update_event_sign_link(
                event_id=request.event_id,
                view_link=comb_result["download_link"],
                column_name="excel_combined"
            )
            
            response_data["appsheet_update_individual"] = appsheet_result_ind
            response_data["appsheet_update_combined"] = appsheet_result_comb
            
            return Response(
                content=json.dumps(response_data),
                media_type="application/json"
            )
        else:
            return Response(
                status_code=500,
                content=json.dumps({"error": "Failed to upload one or both Excel files", "details": response_data}),
                media_type="application/json"
            )
    else:
        # If not uploading, return just one of them as a stream (typically individual)
        # However, since they requested two, it's better to always require drive
        return Response(
            status_code=400,
            content=json.dumps({"error": "upload_to_drive=False is not supported for this multi-file endpoint"}),
            media_type="application/json"
        )

