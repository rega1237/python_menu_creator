import json
import os
from app.services.estimate_docx_generator import EstimateDocxGenerator
from app.schemas.estimate_total import EstimateTotalRequest

# Sample data matching the AppSheet structure I proposed
sample_json = {
  "event_id": "test_id",
  "client": {
    "name": "Simons Foundation",
    "address": "160 5th Ave, New York",
    "email": "info@simonsfoundation.org"
  },
  "client_representative": {
    "name": "Lucas Cacace",
    "email": "lcacace@simonsfoundation.org",
    "formatted_phone": "(646) 599-0962"
  },
  "event": {
    "name": "SIMONS SOCIETY OF FELLOWS DINNER",
    "address": "160 5th Ave #2nd, New York, NY 10010, USA",
    "code": "SSFD-2026",
    "date_formatted": "Wednesday, April 29th 2026",
    "end_date_formatted": "Wednesday, April 29th 2026",
    "guests": 42,
    "dietary_restrictions": "Vegan / Vegetarian options included for all courses."
  },
  "meals": [
    {
      "show_date_header": True,
      "date_header": "Wednesday, April 29th, 2026",
      "category_name": "Hors D'oeuvres",
      "time_range": "7:00 PM to 7:45 PM",
      "description": "Passed selection",
      "provide_by_client": False,
      "category_precio_guest": "Seafood Options",
      "total_category_precio": "$420.00",
      "total_food_por_dia": "$1,200.00",
      "subcategories": [
        {
          "name": "Seafood Options",
          "description": "Premium local catch",
          "items": [
            {
              "name": "Cannoli lobster salad",
              "description": "with spicy red pepper-mayo",
              "diet_options": "GF"
            }
          ]
        },
        {
          "name": "Vegan / Vegetarian Options.",
          "items": [
            {
              "name": "Artichokes & sundried tomatoes mousse",
              "description": "in phyllo cups and basil caviar.",
              "diet_options": "VG"
            }
          ]
        }
      ]
    },
    {
      "category_name": "Plated Dinner",
      "time_range": "7:45 PM to 9:30 PM",
      "description": "Served alongside a selection of artisanal bread and a delightful olive oil infused with a variety of herbs.",
      "category_precio_guest": "Main Course Selection",
      "total_category_precio": "$1,260.00",
      "subcategories": [
        {
          "name": "Appetizer - First Course",
          "items": [
            {
              "name": "Roasted Baby Romaine lettuce",
              "description": "fresh shaved fennel, marinated with honey mustard, fried capers & crumbled Gorgonzola cheese",
              "diet_options": "GF"
            }
          ]
        }
      ]
    }
  ],
  "labor_services": [
    {
      "show_date_header": True,
      "date_header": "Wednesday, April 29th",
      "show_hours_header": True,
      "hours": "5",
      "name": "Wait Staff (4)",
      "total": "$800.00"
    }
  ],
  "extras_events": [
    {
      "show_date_header": True,
      "date_header": "Wednesday, April 29th",
      "is_rental": True,
      "name": "Tableware & Linens",
      "total": "$350.00"
    }
  ],
  "financials": {
    "total_food_service": "$2,500.00",
    "total_labor_cost": "$800.00",
    "total_extras_events": "$350.00",
    "tax_name": "NJ Sales Tax",
    "tax_rate": "6.625%",
    "total_tax": "$12.00",
    "total_extras_sales": "$0",
    "service_charge_rate": "22%",
    "total_service_charge": "$550.00",
    "discount": "0",
    "donation": "0",
    "total_credit_card": "0",
    "gratuity": "0",
    "total_estimate": "$4,212.00"
  }
}

def test_generation():
    generator = EstimateDocxGenerator()
    request = EstimateTotalRequest(**sample_json)
    docx_stream = generator.generate_docx(request)
    
    output_path = "/Users/rafaelguzman/Desktop/proyectos/python_menu_creator/REFACTOR_TEST_RESULT.docx"
    with open(output_path, "wb") as f:
        f.write(docx_stream.read())
    
    print(f"Test complete. Result saved to: {output_path}")

if __name__ == "__main__":
    test_generation()
