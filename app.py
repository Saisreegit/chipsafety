from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import os
import tempfile
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import column_index_from_string

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

excel_data = {}  # Cache data in memory

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]
    if file and file.filename.endswith((".xlsx", ".xls")):
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(filepath)
        xls = pd.ExcelFile(filepath)
        excel_data[file.filename] = {"path": filepath, "sheets": {}}
        return jsonify({"message": "Uploaded", "filename": file.filename, "sheets": xls.sheet_names})
    return jsonify({"error": "Invalid file format"}), 400

@app.route("/edit", methods=["GET"])
def edit():
    filename = request.args.get("filename")
    sheet_name = request.args.get("sheet")
    file_info = excel_data.get(filename)

    if not file_info or not os.path.exists(file_info["path"]):
        return jsonify({"error": "File not found"}), 404

    wb = load_workbook(file_info["path"], data_only=True)
    if sheet_name not in wb.sheetnames:
        return jsonify({"error": "Sheet not found"}), 404

    ws = wb[sheet_name]
    data = list(ws.values)
    
    if not data or not data[0]:
        return jsonify({"error": "No data found"}), 404

    headers = list(data[0])
    rows = data[1:]
    df = pd.DataFrame(rows, columns=headers)
    df.fillna("", inplace=True)

    dropdowns = {}
    for dv in ws.data_validations.dataValidation:
        if dv.formula1 and dv.type == "list":
            try:
                start_cell = dv.sqref.split(":")[0]
                col_letter = ''.join(filter(str.isalpha, start_cell))
                col_idx = column_index_from_string(col_letter) - 1
                if 0 <= col_idx < len(headers):
                    header = headers[col_idx]
                    dropdowns[str(header)] = dv.formula1.replace('"', '').split(',')
            except Exception as e:
                print(f"Dropdown parse error: {e}")

    return jsonify({
        "columns": headers,
        "data": df.to_dict(orient="records"),
        "dropdowns": dropdowns
    })

@app.route("/save", methods=["POST"])
def save():
    data = request.json
    filename = data.get("filename")
    sheet_name = data.get("sheet")
    edited_data = data.get("data")

    if filename not in excel_data:
        return jsonify({"error": "File not found"}), 404

    filepath = excel_data[filename]["path"]
    wb = load_workbook(filepath)
    if sheet_name not in wb.sheetnames:
        return jsonify({"error": "Sheet not found"}), 404

    ws = wb[sheet_name]

    # Get headers from first row
    headers = [cell.value for cell in ws[1]]
    header_index_map = {header: idx + 1 for idx, header in enumerate(headers)}

    # Clear only cell values from row 2 onwards
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    # Write updated data row by row (preserving formatting)
    for row_idx, row_data in enumerate(edited_data, start=2):
        for col_name, value in row_data.items():
            col_idx = header_index_map.get(col_name)
            if col_idx:
                ws.cell(row=row_idx, column=col_idx).value = value

    wb.save(filepath)
    return jsonify({"message": "Saved successfully"})

@app.route("/download", methods=["GET"])
def download():
    filename = request.args.get("filename")
    custom_name = request.args.get("custom_name", "Edited_File.xlsx")
    if filename in excel_data:
        filepath = excel_data[filename]["path"]
        return send_file(filepath, as_attachment=True, download_name=custom_name)
    return jsonify({"error": "File not found"}), 404

if __name__ == "__main__":
    app.run(debug=True)
