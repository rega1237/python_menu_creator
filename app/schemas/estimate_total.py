from pydantic import BaseModel, ConfigDict
from typing import List, Optional

class ClientInfo(BaseModel):
    name: str = ""
    address: str = ""
    email: str = ""

class ClientRepresentative(BaseModel):
    name: str = ""
    email: str = ""
    formatted_phone: str = ""

class EventInfo(BaseModel):
    name: str = ""
    address: str = ""
    code: str = ""
    date_formatted: str = ""
    end_date_formatted: str = ""
    guests: int = 0
    dietary_restrictions: str = ""

class Subcategory(BaseModel):
    name: str = ""
    description: str = ""
    menu_list: str = ""

class Meal(BaseModel):
    show_date_header: bool = False
    date_header: str = ""
    category_name: str = ""
    time_range: str = ""
    description: str = ""
    category_precio_guest: str = ""
    total_category_precio: str = ""
    provide_by_client: bool = False
    total_food_por_dia: str = ""
    subcategories: List[Subcategory] = []
    
    # AppSheet sends true/false as strings sometimes depending on webhook config
    model_config = ConfigDict(coerce_numbers_to_str=True)

class LaborService(BaseModel):
    show_date_header: bool = False
    date_header: str = ""
    show_hours_header: bool = False
    hours: str = ""
    name: str = ""
    total: str = ""

class ExtrasEvent(BaseModel):
    show_date_header: bool = False
    date_header: str = ""
    show_rentals_header: bool = False
    is_rental: bool = False
    name: str = ""
    total: str = ""
    provide_by_client: bool = False

class Financials(BaseModel):
    total_food_service: str = ""
    total_labor_cost: str = ""
    total_extras_events: str = ""
    tax_name: str = ""
    tax_rate: str = ""
    total_tax: str = ""
    total_extras_sales: str = ""
    service_charge_rate: str = ""
    total_service_charge: str = ""
    discount: str = ""
    donation: str = ""
    total_credit_card: str = ""
    gratuity: str = ""
    total_estimate: str = ""

class EstimateTotalRequest(BaseModel):
    event_id: str = ""
    client: ClientInfo
    client_representative: ClientRepresentative
    event: EventInfo
    meals: List[Meal] = []
    labor_services: List[LaborService] = []
    extras_events: List[ExtrasEvent] = []
    financials: Financials
