import pandas as pd
from io import BytesIO
from typing import Tuple
from app.schemas.excel_menu import ExcelMenuRequest

def parse_concatenated_menus(full_text: str):
    """
    Parses a string like "Menu 1 || Description 1 || GF, V , Menu 2 .. || VG"
    If Description is not provided, it falls back to 2 sections.
    Returns a list of tuples: [(Menu Name, Menu Description, Diet Options), ...]
    """
    if not full_text.strip():
        return []
        
    parts = full_text.split("||")
    if not parts or len(parts) == 1:
        cleaned = full_text.strip()
        return [(cleaned, "", "")] if cleaned else []
        
    valid_diet_options = {"GF", "VG", "V", ""}
    results = []
    
    # Pre-process parts. AppSheet might send "Name || Diet" OR "Name || Description || Diet"
    # To handle arbitrary splits correctly, we group the parts first.
    def finalize_menu(name_chunk, desc_chunk, diet_str):
        name = name_chunk.strip()
        # If there's no desc chunk, it implies the original format "Name || Diet" was used.
        desc = desc_chunk.strip() if desc_chunk is not None else ""
        if name:
            results.append((name, desc, diet_str))

    current_menu_name = parts[0].strip()
    current_menu_desc = None
    
    # We iterate over the remaining parts broken by ||
    for i in range(1, len(parts)):
        subparts = parts[i].split(",")
        current_diet_options = []
        next_menu_name_parts = []
        found_next_menu = False
        
        # Determine if this part has valid diet options or if it contains the start of the next menu
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
                
        # Now we decide whether this chunk was a diet string, or a middle description chunk.
        # If found_next_menu is False AND we haven't seen a description yet AND there is another || coming up,
        # it is possible this entire block is a description string.
        # But given the strict comma-split logic, a description without a comma would be swallowed into current_diet_options
        # if it happened to match "GF", "VG", etc.
        # A safer heuristic: If we have a next chunk (i < len(parts) - 1) and NO diet options were found in this block, 
        # or we explicitly define the structure, we can map it.
        # Standard format defined by user: "Name || Menu Description || Diet"
        
        # If there's another "||", the current parts[i] before the comma MIGHT be the description.
        # Let's rebuild the un-split string for this section
        rebuilt_section = parts[i]
        
        # Because AppSheet lists are comma separated, if there's a next menu, it's after a comma.
        if i < len(parts) - 1:
            # We are in the middle chunk. If the original string had 2 "||" for this item, 
            # this chunk is the description.
            if current_menu_desc is None:
                # It's a description chunk. Next menus will be found in later chunks.
                # Careful: The next menu could start in this chunk if it's comma separated.
                if found_next_menu:
                     # e.g., "Desc , NextMenuName". This means the current item ends here without diet options.
                     diet_str = ", ".join(current_diet_options)
                     current_menu_desc = rebuilt_section.split(',')[0].strip() # Simplistic extraction
                     finalize_menu(current_menu_name, current_menu_desc, diet_str)
                     
                     current_menu_name = ", ".join(next_menu_name_parts).strip()
                     current_menu_desc = None
                else:
                     # It's entirely description
                     current_menu_desc = rebuilt_section.strip()
            else:
                # We already have a desc. Find diet options and next menu.
                diet_str = ", ".join(current_diet_options)
                finalize_menu(current_menu_name, current_menu_desc, diet_str)
                current_menu_name = ", ".join(next_menu_name_parts).strip()
                current_menu_desc = None
        else:
            # Final chunk
            diet_str = ", ".join(current_diet_options)
            finalize_menu(current_menu_name, current_menu_desc, diet_str)
            current_menu_name = ", ".join(next_menu_name_parts).strip()
            # If there's dangling text after the last diet options (e.g. valid menu without || at the end), add it
            if current_menu_name:
                finalize_menu(current_menu_name, None, "")
        
    return results

def sort_dataframe_by_date(df: pd.DataFrame) -> pd.DataFrame:
    """Helper to sort dataframe by Date."""
    if df.empty:
        return df
    
    df['_sort_date'] = pd.to_datetime(df['Date'], errors='coerce')
    df = df.sort_values(by=['_sort_date'], na_position='first')
    return df.drop(columns=['_sort_date'])

def sort_dataframe_by_date(df: pd.DataFrame) -> pd.DataFrame:
    """Helper to sort dataframe by Date."""
    if df.empty:
        return df
    
    # Safely convert to datetime/time for sorting without overwriting original string format
    df['_sort_date'] = pd.to_datetime(df['Date'], errors='coerce')
    
    # Sort placing NaT (empty times) first
    df = df.sort_values(by=['_sort_date'], na_position='first')
        
    return df.drop(columns=['_sort_date'])

def generate_individual_excel(request: ExcelMenuRequest) -> BytesIO:
    """
    Generates an Excel where each individual menu item is a row.
    Columns: Date | Category | Category Desc | Subcategory | Menu | Description | Diet Options
    """
    rows = []
    
    for meal in request.all_meals:
        base_data = {
            "Date": meal.date,
            "Category": meal.category,
            "Category Desc": meal.description # Changed from just "Description" to avoid name collision
        }
        
        for item in meal.items:
            # Skip empty entries if desired
            if not item.subcat.strip() and not item.menu.strip():
                continue
                
            # Use the new robust parsing function
            parsed_menus = parse_concatenated_menus(item.menu)
            
            for menu_name, menu_description, diet_options in parsed_menus:
                row = base_data.copy()
                row["Subcategory"] = item.subcat
                row["Menu"] = menu_name
                row["Description"] = menu_description
                row["Diet Options"] = diet_options
                
                rows.append(row)
                
    df = pd.DataFrame(rows)
    df = sort_dataframe_by_date(df)
    
    # Important: Reorder columns to ensure "Description" is between "Menu" and "Diet Options"
    # and parent Category description is clearly differentiated.
    column_order = ["Date", "Category", "Category Desc", "Subcategory", "Menu", "Description", "Diet Options"]
    # Add any missing columns safely
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    df = df[column_order]
    
    # Save to memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Individual Menus')
    
    output.seek(0)
    return output

def generate_combined_excel(request: ExcelMenuRequest) -> BytesIO:
    """
    Generates an Excel where each row is a Meal's subcategory.
    Columns: Date | Category | Category Desc | Subcategory | Menu
    """
    rows = []
    
    for meal in request.all_meals:
        base_data = {
            "Date": meal.date,
            "Category": meal.category,
            "Category Desc": meal.description
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
            for menu_name, menu_desc, diet_options in parsed_menus:
                
                base_text = menu_name
                
                if diet_options:
                    formatted_menus.append(f"{base_text} || {diet_options}")
                else:
                    # Omit the || when there are no diet options
                    formatted_menus.append(base_text)
                    
            row["Menu"] = " , ".join(formatted_menus)
            
            rows.append(row)
        
    df = pd.DataFrame(rows)
    df = sort_dataframe_by_date(df)
    
    column_order = ["Date", "Category", "Category Desc", "Subcategory", "Menu"]
    for col in column_order:
        if col not in df.columns:
            df[col] = ""
    df = df[column_order]
    
    # Save to memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Combined Menus')
    
    output.seek(0)
    return output
