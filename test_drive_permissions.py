import os
import json
from io import BytesIO
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload

# Using the broader scope I just added in the service
SCOPES = ['https://www.googleapis.com/auth/drive']

def test_drive_permissions():
    # ID real detectado por el script (¡Ojo con la 'x' minúscula!)
    folder_id = "1kUQU9tf9rxNfEbxppGaJFuAq8hFLmsyg"
    # Autodiscover the JSON in the current directory if it starts with 'menu-appsheet'
    creds_files = [f for f in os.listdir('.') if f.startswith('menu-appsheet') and f.endswith('.json')]
    if creds_files:
        creds_path = os.path.abspath(creds_files[0])
    else:
        creds_path = "/Users/rafaelguzman/Desktop/proyectos/python_menu_creator/menu-appsheet-4b3313740817.json"
    
    print(f"--- Probando conexión con Folder ID: {folder_id} ---")
    print(f"Usando credenciales de: {creds_path}")
    
    if not os.path.exists(creds_path):
        print(f"Error: No se encontró el JSON en {creds_path}")
        return

    try:
        creds = service_account.Credentials.from_service_account_file(creds_path, scopes=SCOPES)
        print(f"Identidad de la cuenta: {creds.service_account_email}")
        service = build('drive', 'v3', credentials=creds)
        
        print("\n--- Listando archivos accesibles por esta cuenta ---")
        results = service.files().list(
            pageSize=10, fields="nextPageToken, files(id, name)",
            supportsAllDrives=True, includeItemsFromAllDrives=True
        ).execute()
        items = results.get('files', [])
        if not items:
            print("No se encontraron archivos accesibles. (¿Está compartida la carpeta con este email?)")
        else:
            for item in items:
                print(f"- {item['name']} ({item['id']})")
        
        # Intentar obtener info de la carpeta con soporte para Shared Drives
        print(f"\nVerificando existencia de la carpeta {folder_id}...")
        try:
            folder = service.files().get(
                fileId=folder_id, 
                fields='id, name',
                supportsAllDrives=True
            ).execute()
            print(f"✅ Carpeta encontrada: {folder.get('name')} (ID: {folder.get('id')})")
        except Exception as e_inner:
            print(f"❌ No se pudo encontrar la carpeta específica: {e_inner}")
        
    except Exception as e:
        print(f"❌ Falló la prueba general: {e}")

if __name__ == "__main__":
    test_drive_permissions()
