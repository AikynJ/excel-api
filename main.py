from flask import Flask, request, send_file
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import io

app = Flask(__name__)

@app.route("/generate-xlsx", methods=["POST"])
def generate_xlsx():
    data = request.json.get("data", [])
    df = pd.DataFrame(data)

    column_order = [
        "НАИМЕНОВАНИЕ", "АРТИКУЛ", "БАРКОД", "КОЛ-ВО", "V_Размер",
        "ЦЕНА ПОСТАВКИ (KZT)", "РОЗНИЧНАЯ ЦЕНА (KZT)", "КАТЕГОРИЯ",
        "БРЕНД", "ЕДИНИЦА ИЗМЕРЕНИЯ", "ПОСТАВЩИК", "V_Цвет",
        "mark_code", "ЦЕНА В USD"
    ]
    df = df[[col for col in column_order if col in df.columns]]

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')

    output.seek(0)
    wb = load_workbook(output)
    ws = wb.active

    # Fonts
    font_header_red = Font(name='Calibri', size=11, bold=True, color='FF0000')
    font_header_black = Font(name='Calibri', size=11, bold=True)
    font_gothic_9 = Font(name='Century Gothic', size=9)
    font_calibri_10 = Font(name='Calibri', size=10)
    font_calibri_11 = Font(name='Calibri', size=11)
    font_arial_9 = Font(name='Arial', size=9)
    font_times_9 = Font(name='Times New Roman', size=9)

    # Fills
    fill_qty = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
    fill_barcode = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')

    # Borders
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    # Alignments
    align_center = Alignment(horizontal='center', vertical='center')
    align_right = Alignment(horizontal='right', vertical='center')
    align_left = Alignment(horizontal='left', vertical='center')

    # Apply header styles
    for idx, cell in enumerate(ws[1], 1):
        col_name = cell.value
        if col_name in ["НАИМЕНОВАНИЕ", "АРТИКУЛ", "БАРКОД", "КОЛ-ВО", "ЦЕНА ПОСТАВКИ (KZT)", "РОЗНИЧНАЯ ЦЕНА (KZT)"]:
            cell.font = font_header_red
        else:
            cell.font = font_header_black
        cell.alignment = align_center
        cell.border = thin_border

    headers = [cell.value for cell in ws[1]]

    # Apply data styles
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
        for cell in row:
            header = ws.cell(row=1, column=cell.col_idx).value
            if header == "НАИМЕНОВАНИЕ":
                cell.font = font_gothic_9
                cell.border = thin_border
            elif header == "АРТИКУЛ":
                cell.font = font_calibri_10
            elif header == "БАРКОД":
                cell.font = font_calibri_11
                cell.fill = fill_barcode
                cell.alignment = align_center
            elif header == "КОЛ-ВО":
                cell.font = font_gothic_9
                cell.fill = fill_qty
                cell.border = thin_border
                cell.alignment = align_center
            elif header == "V_Размер":
                cell.font = font_calibri_11
            elif header == "ЦЕНА ПОСТАВКИ (KZT)":
                cell.font = font_arial_9
                cell.number_format = "0.00"
                cell.alignment = align_right
            elif header == "РОЗНИЧНАЯ ЦЕНА (KZT)":
                cell.font = font_calibri_11
                cell.number_format = "0.00"
                cell.alignment = align_right
            elif header == "КАТЕГОРИЯ":
                cell.font = font_calibri_11
            elif header in ["БРЕНД", "ЕДИНИЦА ИЗМЕРЕНИЯ", "ПОСТАВЩИК"]:
                cell.font = font_calibri_11
            elif header == "V_Цвет":
                cell.font = font_gothic_9
                cell.border = thin_border
            elif header == "mark_code":
                cell.font = font_calibri_11
            elif header == "ЦЕНА В USD":
                cell.font = font_times_9
                cell.border = thin_border
                cell.number_format = "0.00"
                cell.alignment = align_right

    # Auto-fit column width
    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max_length + 2

    final_output = io.BytesIO()
    wb.save(final_output)
    final_output.seek(0)

    return send_file(final_output, download_name="formatted.xlsx", as_attachment=True)
