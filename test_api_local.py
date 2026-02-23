import requests
import json
import os

def test_local_api():
    url = "http://127.0.0.1:8000/api/v1/menus/generate"
    json_path = "sample_menu.json"
    output_path = "resultado_test.docx"

    if not os.path.exists(json_path):
        print(f"Error: No se encontró el archivo {json_path}")
        return

    print(f"Enviando petición a {url}...")
    
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    try:
        response = requests.post(url, json=data)
        
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                f.write(response.content)
            print(f"✅ ¡Éxito! El archivo se ha guardado como: {os.path.abspath(output_path)}")
        else:
            print(f"❌ Error {response.status_code}: {response.text}")
            
    except requests.exceptions.ConnectionError:
        print("❌ Error: No se pudo conectar con el servidor. Asegúrate de que 'uvicorn app.main:app --reload' esté corriendo.")

if __name__ == "__main__":
    test_local_api()
