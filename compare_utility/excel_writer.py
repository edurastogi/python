from openpyxl import Workbook
from openpyxl.styles import PatternFill
from openpyxl.cell import WriteOnlyCell
import os

EXCEL_MAX_ROWS = 1048576  # Excel .xlsx row limit per sheet

def append_row_with_split(ws_list, row, row_counter, base_title, header):
    ws, count = ws_list[-1]
    if count >= EXCEL_MAX_ROWS:
        idx = len(ws_list) + 1
        ws_new = ws.parent.create_sheet(f"{base_title}_{idx}")
        ws_new.append(header)
        ws_list.append((ws_new, 1))
        ws, count = ws_list[-1]
    ws.append(row)
    ws_list[-1] = (ws, count + 1)

# Optionally, you can add more Excel writing helpers here as needed.

