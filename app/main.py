from fastapi import FastAPI, Response, Query
from app.schemas.menu import MenuRequest
from app.services.docx_generator import generate_menu_docx
from app.services.google_drive_service import drive_service

app = FastAPI(title="Menu Creator Service")

@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.post("/api/v1/menus/generate", response_class=Response)
async def generate_menu(
    request: MenuRequest, 
    upload_to_drive: bool = Query(False, description="Si es True, sube el archivo a Google Drive y devuelve JSON con el link.")
):
    docx_stream = generate_menu_docx(request)
    
    filename = f"menu_{request.all_meals[0].categoria}_{request.all_meals[0].fecha}.docx".replace(" ", "_")
    
    if upload_to_drive:
        result = drive_service.upload_file(docx_stream, filename)
        if result["success"]:
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
