import re
from pydantic import BaseModel, ConfigDict, field_validator
from typing import List, Optional

def format_to_us_date(val: str) -> str:
    """Converts DD/MM/YYYY or DD/MM/YY to MM/DD/YYYY or MM/DD/YY (US format)."""
    if not val:
        return val
    val_clean = str(val).strip()
    
    # 1. Match DD/MM/YYYY
    match1 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{4})$", val_clean)
    if match1:
        day, month, year = match1.groups()
        return f"{month.zfill(2)}/{day.zfill(2)}/{year}"
        
    # 2. Match DD/MM/YY
    match2 = re.match(r"^(\d{1,2})/(\d{1,2})/(\d{2})$", val_clean)
    if match2:
        day, month, year = match2.groups()
        return f"{month.zfill(2)}/{day.zfill(2)}/{year}"
        
    return val

def format_time_range(val: str) -> str:
    """Injects correct spaces around 'to' in meal time ranges."""
    if not val:
        return val
    # Replace "to" with " to " only if it is preceded by a digit or M/m 
    # and followed by a digit or A/a/P/p to avoid modifying other words.
    res = re.sub(r'([0-9Mm])\s*to\s*([0-9AaPp])', r'\1 to \2', str(val), flags=re.IGNORECASE)
    return ' '.join(res.split())


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

    @field_validator('date_formatted', 'end_date_formatted', mode='after')
    @classmethod
    def format_dates(cls, v):
        return format_to_us_date(v)

class MenuItem(BaseSchema):
    name: str = ""
    description: str = ""
    diet_options: str = ""


class Meal(BaseSchema):
    show_date_header: bool = False
    order: int = 0
    date_header: str = ""
    category_name: str = ""
    time_range: str = ""
    description: str = ""
    category_precio_guest: str = ""
    total_category_precio: str = ""
    provide_by_client: bool = False
    total_food_por_dia: str = ""
    
    # Fields for Per Day Estimate
    show_date_header_2: bool = False
    date_day_name: str = ""
    guest_count: str = ""
    show_guest_header: bool = False
    total_category_precio_guest_por_dia: str = ""
    
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

    @field_validator('date_header', mode='after')
    @classmethod
    def format_date_header(cls, v):
        return format_to_us_date(v)

    @field_validator('time_range', mode='after')
    @classmethod
    def format_times(cls, v):
        return format_time_range(v)

class LaborService(BaseSchema):
    show_date_header: bool = False
    order: int = 0
    date_header: str = ""
    show_hours_header: bool = False
    hours: str = ""
    name: str = ""
    total: str = ""

    @field_validator('date_header', mode='after')
    @classmethod
    def format_date_header(cls, v):
        return format_to_us_date(v)

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

    @field_validator('date_header', mode='after')
    @classmethod
    def format_date_header(cls, v):
        return format_to_us_date(v)

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
    credit_card_percent: str = "0"
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
