import os
import json
from google_auth_oauthlib.flow import InstalledAppFlow

# Scopes required for Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']

def get_refresh_token():
    print("--- Generador de Refresh Token de Google ---")
    print("Para usar esto, primero debes crear 'Credenciales de OAuth' tipo 'App de escritorio' en Google Cloud Console.")
    
    client_id = input("Introduce tu GOOGLE_CLIENT_ID: ").strip()
    client_secret = input("Introduce tu GOOGLE_CLIENT_SECRET: ").strip()
    
    client_config = {
        "installed": {
            "client_id": client_id,
            "client_secret": client_secret,
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    }

    flow = InstalledAppFlow.from_client_config(client_config, SCOPES)
    # Usamos el puerto 8080 fijo para que sea fácil registrarlo si es necesario
    creds = flow.run_local_server(port=8080)

    print("\n--- ¡ÉXITO! Copia estos valores para Render ---\n")
    print(f"GOOGLE_CLIENT_ID: {client_id}")
    print(f"GOOGLE_CLIENT_SECRET: {client_secret}")
    print(f"GOOGLE_REFRESH_TOKEN: {creds.refresh_token}")
    print("\n----------------------------------------------")

if __name__ == "__main__":
    get_refresh_token()
