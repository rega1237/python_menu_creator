from fastapi import FastAPI, Response
from app.schemas.menu import MenuRequest
from app.services.docx_generator import generate_menu_docx

app = FastAPI(title="Menu Creator Service")

@app.get("/")
async def root():
    return {"message": "Hello World"}

@app.post("/api/v1/menus/generate", response_class=Response)
async def generate_menu(request: MenuRequest):
    docx_stream = generate_menu_docx(request)
    
    headers = {
        'Content-Disposition': 'attachment; filename="menu_generado.docx"'
    }
    
    return Response(
        content=docx_stream.read(),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers=headers
    )
