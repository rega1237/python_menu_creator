from fastapi.testclient import TestClient
from app.main import app

client = TestClient(app)

def test_read_main():
    response = client.get("/")
    assert response.status_code == 200
    assert response.json() == {"message": "Hello World"}

def test_generate_menu_endpoint():
    payload = {
        "all_meals": [
            {
                "categoria": "Desayuno",
                "fecha": "2023-11-01",
                "descripcion": "Desayuno completo",
                "items": [
                    {"subcat": "Frutas", "menu": "Manzana, Pera"},
                    {"subcat": "Bebidas", "menu": "Café, Jugo"}
                ]
            }
        ]
    }
    
    response = client.post("/api/v1/menus/generate", json=payload)
    
    assert response.status_code == 200
    assert response.headers["content-type"] == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    assert "attachment; filename" in response.headers["content-disposition"]
    
    # Check that we actually got some binary content back
    assert len(response.content) > 1000  # A basic empty docx is typically ~11KB
