from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Load the styled Excel file
filepath = "/mnt/data/styled_output_final.xlsx"
wb = load_workbook(filepath)
ws = wb.active

# Auto-fit all columns by adjusting width based on max length of cell content
for col in ws.columns:
    max_length = 0
    column = col[0].column  # Get column number
    for cell in col:
        try:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        except:
            pass
    adjusted_width = max_length + 2
    ws.column_dimensions[get_column_letter(column)].width = adjusted_width

# Save updated file
updated_path = "/mnt/data/styled_output_autofit.xlsx"
wb.save(updated_path)

updated_path
