import os
import requests
import logging

logger = logging.getLogger(__name__)

class AppSheetService:
    def __init__(self):
        self.app_id = os.getenv("APPSHEET_APP_ID", "").strip()
        self.access_key = os.getenv("APPSHEET_ACCESS_KEY", "").strip()
        self.table_name = "BDEvents"
        
    def update_event_sign_link(self, event_id: str, view_link: str) -> dict:
        """Updates the SINGS_GENERAL_WORD column in AppSheet for the given event_id."""
        if not (self.app_id and self.access_key):
            logger.error("AppSheet credentials missing")
            return {"success": False, "error": "AppSheet credentials missing"}
            
        url = f"https://api.appsheet.com/api/v1/apps/{self.app_id}/tables/{self.table_name}/Action"
        
        headers = {
            'ApplicationAccessKey': self.access_key,
            'Content-Type': 'application/json'
        }
        
        payload = {
            "Action": "Edit",
            "Properties": {
                "Locale": "en-US",
                "Timezone": "Eastern Standard Time"
            },
            "Rows": [
                {
                    "ID": event_id,
                    "SINGS_GENERAL_WORD": view_link
                }
            ]
        }
        
        try:
            logger.info(f"Sending callback to AppSheet for event_id: {event_id}")
            logger.info(f"AppSheet Payload: {payload}")
            response = requests.post(url, headers=headers, json=payload)
            response.raise_for_status()
            
            result = response.json()
            logger.info(f"AppSheet API Response: {result}")
            
            return {"success": True, "result": result}
            
        except Exception as e:
            logger.error(f"Error calling AppSheet API: {e}")
            if hasattr(e, 'response') and e.response:
                logger.error(f"Response content: {e.response.text}")
            return {"success": False, "error": str(e)}

# Singleton instance
appsheet_service = AppSheetService()
