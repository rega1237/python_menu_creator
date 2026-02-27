import json
from app.schemas.estimate_total import EstimateTotalRequest
from app.services.estimate_docx_generator import EstimateDocxGenerator

test_payload = {
  "event_id": "ABC-123",
  "client": {
    "name": "Rafael Guzman",
    "address": "Santo Domingo, DR",
    "email": "rafael@example.com"
  },
  "client_representative": {
    "name": "Jane Doe",
    "email": "jane@business.com",
    "formatted_phone": "+1 555-0198"
  },
  "event": {
    "name": "Annual Corporate Summit 2025",
    "address": "789 Event Center Blvd, NY",
    "code": "EV-2025-099",
    "date_formatted": "Saturday, October 14th 2025",
    "end_date_formatted": "Sunday, October 15th 2025",
    "guests": 200,
    "dietary_restrictions": "Vegan, Gluten-Free, Nut Allergy"
  },
  "meals": [
    {
      "show_date_header": True,
      "date_header": "Saturday, October 14th",
      "category_name": "Breakfast",
      "time_range": "8:00 AM to 10:00 AM",
      "description": "Premium Continental Breakfast",
      "category_precio_guest": "$25.00/person",
      "total_category_precio": "$5,000.00",
      "provide_by_client": False,
      "total_food_por_dia": "$5,000.00",
      "subcategories": [
        {
          "name": "Pastries",
          "description": "Assorted freshly baked goods",
          "menu_list": "Croissants\nMuffins\nGluten-Free Scones"
        },
        {
          "name": "Beverages",
          "description": "Hot and cold drinks",
          "menu_list": "Coffee\nTea\nFresh Orange Juice"
        }
      ]
    }
  ],
  "labor_services": [
    {
      "show_date_header": True,
      "date_header": "Saturday, October 14th",
      "show_hours_header": True,
      "hours": "8",
      "name": "1 Event Captain, 4 Servers, 1 Bartender",
      "total": "$1,400.00"
    }
  ],
  "extras_events": [
    {
      "show_date_header": True,
      "date_header": "Saturday, October 14th",
      "show_rentals_header": True,
      "is_rental": True,
      "name": "Luxury Table Linens & Chairs",
      "total": "$800.00",
      "provide_by_client": False
    },
    {
      "show_date_header": False,
      "date_header": "",
      "show_rentals_header": True,
      "is_rental": False,
      "name": "Custom Floral Centerpieces",
      "total": "$450.00",
      "provide_by_client": False
    }
  ],
  "financials": {
    "total_food_service": "$5,000.00",
    "total_labor_cost": "$1,400.00",
    "total_extras_events": "$1,250.00",
    "tax_name": "NY State Tax",
    "tax_rate": "8.875%",
    "total_tax": "$678.94",
    "total_extras_sales": "$450.00",
    "service_charge_rate": "18%",
    "total_service_charge": "$1,377.00",
    "discount": "$200.00",
    "donation": "$0.00",
    "total_credit_card": "$0.00",
    "gratuity": "$500.00",
    "total_estimate": "$10,005.94"
  }
}

def main():
    request_data = EstimateTotalRequest(**test_payload)
    generator = EstimateDocxGenerator()
    docx_bytes = generator.generate_docx(request_data)
    
    with open("test_estimate.docx", "wb") as f:
        f.write(docx_bytes.getbuffer())
        print("Generated test_estimate.docx successfully.")

if __name__ == "__main__":
    main()
