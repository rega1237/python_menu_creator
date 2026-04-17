from pydantic import BaseModel, ConfigDict
from typing import List, Optional

class BaseSchema(BaseModel):
    model_config = ConfigDict(coerce_numbers_to_str=True)

class ClientInfo(BaseSchema):
    name: str = ""
    address: str = ""
    email: str = ""

class ClientRepresentative(BaseSchema):
    name: str = ""
    email: str = ""
    formatted_phone: str = ""

class EventInfo(BaseSchema):
    name: str = ""
    address: str = ""
    code: str = ""
    date_formatted: str = ""
    end_date_formatted: str = ""
    guests: int = 0
    dietary_restrictions: str = ""

class MenuItem(BaseSchema):
    name: str = ""
    description: str = ""
    diet_options: str = ""


class Meal(BaseSchema):
    show_date_header: bool = False
    date_header: str = ""
    category_name: str = ""
    time_range: str = ""
    description: str = ""
    category_precio_guest: str = ""
    total_category_precio: str = ""
    provide_by_client: bool = False
    total_food_por_dia: str = ""
    
    # Flattened subcategories to match AppSheet fixed columns
    subcategory_1_name: Optional[str] = ""
    subcategory_1_description: Optional[str] = ""
    subcategory_1_items: List[MenuItem] = []
    
    subcategory_2_name: Optional[str] = ""
    subcategory_2_description: Optional[str] = ""
    subcategory_2_items: List[MenuItem] = []
    
    subcategory_3_name: Optional[str] = ""
    subcategory_3_description: Optional[str] = ""
    subcategory_3_items: List[MenuItem] = []
    
    subcategory_4_name: Optional[str] = ""
    subcategory_4_description: Optional[str] = ""
    subcategory_4_items: List[MenuItem] = []
    
    subcategory_5_name: Optional[str] = ""
    subcategory_5_description: Optional[str] = ""
    subcategory_5_items: List[MenuItem] = []
    
    subcategory_6_name: Optional[str] = ""
    subcategory_6_description: Optional[str] = ""
    subcategory_6_items: List[MenuItem] = []
    
    subcategory_7_name: Optional[str] = ""
    subcategory_7_description: Optional[str] = ""
    subcategory_7_items: List[MenuItem] = []
    
    subcategory_8_name: Optional[str] = ""
    subcategory_8_description: Optional[str] = ""
    subcategory_8_items: List[MenuItem] = []
    
    subcategory_9_name: Optional[str] = ""
    subcategory_9_description: Optional[str] = ""
    subcategory_9_items: List[MenuItem] = []
    
    subcategory_10_name: Optional[str] = ""
    subcategory_10_description: Optional[str] = ""
    subcategory_10_items: List[MenuItem] = []
    
    subcategory_11_name: Optional[str] = ""
    subcategory_11_description: Optional[str] = ""
    subcategory_11_items: List[MenuItem] = []
    
    subcategory_12_name: Optional[str] = ""
    subcategory_12_description: Optional[str] = ""
    subcategory_12_items: List[MenuItem] = []

class LaborService(BaseSchema):
    show_date_header: bool = False
    date_header: str = ""
    show_hours_header: bool = False
    hours: str = ""
    name: str = ""
    total: str = ""

class ExtrasEvent(BaseSchema):
    show_date_header: bool = False
    date_header: str = ""
    is_rental: bool = False
    is_sales: bool = False
    name: str = ""
    name_rental: str = ""
    name_sales: str = ""
    total: str = ""
    provide_by_client: bool = False

class Financials(BaseSchema):
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

class EstimateTotalRequest(BaseSchema):
    event_id: str = ""
    client: ClientInfo
    client_representative: ClientRepresentative
    event: EventInfo
    meals: List[Meal] = []
    labor_services: List[LaborService] = []
    extras_events: List[ExtrasEvent] = []
    financials: Financials
