from openpyxl.styles import PatternFill

EXCEL_MAX_ROWS = 1048576  # Excel .xlsx row limit per sheet

def get_next_sheet_name(base_title, idx):
    """
    Generate a new sheet name with a suffix if needed.
    """
    return f"{base_title}_{idx}" if idx > 1 else base_title

def append_row_with_split(ws_list, row, row_counter, base_title, header):
    """
    Append a row to the current worksheet, splitting to a new sheet if needed.
    """
    ws, count = ws_list[-1]
    if count >= EXCEL_MAX_ROWS:
        idx = len(ws_list) + 1
        ws_new = ws.parent.create_sheet(get_next_sheet_name(base_title, idx))
        ws_new.append(header)
        ws_list.append((ws_new, 1))
        ws, count = ws_list[-1]
    ws.append(row)
    ws_list[-1] = (ws, count + 1)

# Optionally, you can add more Excel writing helpers here as needed.
