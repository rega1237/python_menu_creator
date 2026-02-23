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
    all_meals: List[MenuData]
