import pytest
from pydantic import ValidationError
from app.schemas.menu import MenuRequest, MenuData, SubCategoryItem

def test_valid_menu_request():
    payload = {
        "all_meals": [
            {
                "categoria": "Desayuno",
                "fecha": "2023-10-27",
                "descripcion": "Desayuno nutritivo",
                "items": [
                    {"subcat": "Bebidas", "menu": "Jugo de Naranja"},
                    {"subcat": "", "menu": ""}
                ]
            }
        ]
    }
    
    # Should not raise exception
    request = MenuRequest(**payload)
    assert len(request.all_meals) == 1
    meal = request.all_meals[0]
    assert meal.categoria == "Desayuno"
    assert meal.fecha == "2023-10-27"
    assert meal.descripcion == "Desayuno nutritivo"
    assert len(meal.items) == 2
    assert meal.items[0].subcat == "Bebidas"
    assert meal.items[0].menu == "Jugo de Naranja"
    assert meal.items[1].subcat == ""

def test_invalid_menu_request_missing_categoria():
    payload = {
        "all_meals": [
            {
                "fecha": "2023-10-27",
                "descripcion": "Desayuno nutritivo",
                "items": []
            }
        ]
    }
    with pytest.raises(ValidationError):
        MenuRequest(**payload)
