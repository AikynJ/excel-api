from flask import Flask, request, send_file
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
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

        header_font = Font(bold=True, color='FF0000')
        green_fill = PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        lightgreen_fill = PatternFill(start_color='CCFFCC', end_color='CCFFCC', fill_type='solid')
        center = Alignment(horizontal='center')
        right = Alignment(horizontal='right')

        for cell in ws[1]:  # Заголовки
            cell.font = header_font
            cell.alignment = center

        for row in ws.iter_rows(min_row=2):
            for cell in row:
                header = ws.cell(row=1, column=cell.col_idx).value
                if header == "КОЛ-ВО":
                    cell.fill = green_fill
                    cell.alignment = center
                elif header == "БАРКОД":
                    cell.fill = lightgreen_fill
                    cell.alignment = center
                elif header in ["ЦЕНА В USD", "ЦЕНА ПОСТАВКИ (KZT)", "РОЗНИЧНАЯ ЦЕНА (KZT)"]:
                    cell.number_format = "0.00"
                    cell.alignment = right

    output.seek(0)
    return send_file(output, download_name="formatted.xlsx", as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
