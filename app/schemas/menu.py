from pydantic import BaseModel, Field
from typing import List

class SubCategoryItem(BaseModel):
    subcat: str
    menu: str

class MenuData(BaseModel):
    categoria: str
    fecha: str
    descripcion: str
    items: List[SubCategoryItem]

class MenuRequest(BaseModel):
    event_id: str = "unknown"
    event_name: str = "event"
    all_meals: List[MenuData]
