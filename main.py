from flask import Flask, request, send_file
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io

app = Flask(__name__)

@app.route("/generate-xlsx", methods=["POST"])
def generate_xlsx():
    data = request.json.get("data", [])
    df = pd.DataFrame(data)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        ws = writer.sheets['Sheet1']

        # Стили
        header_font = Font(bold=True, name='Arial', size=10, color='FF0000')
        green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        lightgreen_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
        center = Alignment(horizontal='center', vertical='center')
        right = Alignment(horizontal='right', vertical='center')
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )

        # Заголовки — стиль
        for cell in ws[1]:
            cell.font = header_font
            cell.alignment = center
            cell.border = thin_border

        headers = [cell.value for cell in ws[1]]

        # Применение стилей к строкам
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(headers)):
            for cell in row:
                col_header = headers[cell.col_idx - 1]
                cell.border = thin_border
                if col_header == "КОЛ-ВО":
                    cell.fill = green_fill
                    cell.alignment = center
                elif col_header == "БАРКОД":
                    cell.fill = lightgreen_fill
                    cell.alignment = center
                elif col_header in ["ЦЕНА В USD", "ЦЕНА ПОСТАВКИ (KZT)", "РОЗНИЧНАЯ ЦЕНА (KZT)"]:
                    cell.number_format = "0.00"
                    cell.alignment = right
                else:
                    cell.alignment = Alignment(horizontal='left', vertical='center')

    output.seek(0)
    return send_file(output, download_name="formatted.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
