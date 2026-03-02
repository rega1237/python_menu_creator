from pydantic import BaseModel
from typing import List, Optional

class ExcelMenuPair(BaseModel):
    subcat: str
    menu: str

class ExcelMealData(BaseModel):
    date: str
    category: str
    description: str
    items: List[ExcelMenuPair]

class ExcelMenuRequest(BaseModel):
    event_id: str
    event_name: str
    all_meals: List[ExcelMealData]
