import pandas as pd
from io import BytesIO
from typing import Tuple
from app.schemas.excel_menu import ExcelMenuRequest

def parse_concatenated_menus(full_text: str):
    """
    Parses a string like "Menu 1 || GF, V , Menu 2, with comma || VG"
    Returns a list of tuples: [(Menu Name, Diet Options), ...]
    """
    if not full_text.strip():
        return []
        
    parts = full_text.split("||")
    if not parts or len(parts) == 1:
        cleaned = full_text.strip()
        return [(cleaned, "")] if cleaned else []
        
    valid_diet_options = {"GF", "VG", "V", ""}
    results = []
    current_menu_name = parts[0].strip()
    
    for i in range(1, len(parts)):
        subparts = parts[i].split(",")
        current_diet_options = []
        next_menu_name_parts = []
        found_next_menu = False
        
        for subpart in subparts:
            cleaned = subpart.strip()
            
            if not found_next_menu:
                if cleaned in valid_diet_options:
                    if cleaned:
                        current_diet_options.append(cleaned)
                else:
                    found_next_menu = True
                    next_menu_name_parts.append(cleaned)
            else:
                next_menu_name_parts.append(cleaned)
                
        diet_str = ", ".join(current_diet_options)
        if current_menu_name:
            results.append((current_menu_name, diet_str))
        
        current_menu_name = ", ".join(next_menu_name_parts).strip()
        
    return results

def generate_individual_excel(request: ExcelMenuRequest) -> BytesIO:
    """
    Generates an Excel where each individual menu item is a row.
    Columns: Date | Clock In | Clock Out | Category | Description | Subcategory | Menu | Diet Options
    """
    rows = []
    
    for meal in request.all_meals:
        # Base row data that is the same for every menu item in this meal
        base_data = {
            "Date": meal.date,
            "Clock In": meal.clock_in,
            "Clock Out": meal.clock_out,
            "Category": meal.category,
            "Description": meal.description
        }
        
        for item in meal.items:
            # Skip empty entries if desired
            if not item.subcat.strip() and not item.menu.strip():
                continue
                
            # Use the new robust parsing function
            parsed_menus = parse_concatenated_menus(item.menu)
            
            for menu_name, diet_options in parsed_menus:
                row = base_data.copy()
                row["Subcategory"] = item.subcat
                row["Menu"] = menu_name
                row["Diet Options"] = diet_options
                
                rows.append(row)
                
    df = pd.DataFrame(rows)
    
    # Save to memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Individual Menus')
    
    output.seek(0)
    return output

def generate_combined_excel(request: ExcelMenuRequest) -> BytesIO:
    """
    Generates an Excel where each row is a Meal's subcategory.
    Columns: Date | Clock In | Clock Out | Category | Description | Subcategory | Menu
    """
    rows = []
    
    for meal in request.all_meals:
        base_data = {
            "Date": meal.date,
            "Clock In": meal.clock_in,
            "Clock Out": meal.clock_out,
            "Category": meal.category,
            "Description": meal.description
        }
        
        for item in meal.items:
            # Skip empty entries if desired
            if not item.subcat.strip() and not item.menu.strip():
                continue
            
            row = base_data.copy()
            row["Subcategory"] = item.subcat
            
            # Use the robust parsed menus to reformat the string gracefully
            parsed_menus = parse_concatenated_menus(item.menu)
            
            formatted_menus = []
            for menu_name, diet_options in parsed_menus:
                if diet_options:
                    formatted_menus.append(f"{menu_name} || {diet_options}")
                else:
                    # Omit the || when there are no diet options
                    formatted_menus.append(menu_name)
                    
            row["Menu"] = " , ".join(formatted_menus)
            
            rows.append(row)
        
    df = pd.DataFrame(rows)
    
    # Save to memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Menus')
    
    output.seek(0)
    return output
