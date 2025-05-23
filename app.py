from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import os
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries
from db import insert_excel_data  # make sure this works

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = os.path.join(os.getcwd(), "uploads")
os.makedirs(UPLOAD_FOLDER, exist_ok=True)  # Ensure folder exists
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

excel_data = {}  # Cache for uploaded Excel files info

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    # Expecting file input name to be "file" to keep consistent with frontend
    file = request.files.get("file")
    if file and file.filename.endswith((".xlsx", ".xls")):
        filepath = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(filepath)

        try:
            wb = load_workbook(filepath)
            sheet = wb.active
            # Cache file info for later use
            excel_data[file.filename] = {"path": filepath}

            # Return some basic info for confirmation and frontend use
            return jsonify({
                "message": f"Uploaded: {file.filename}",
                "filename": file.filename,
                "sheets": sheet_names
            })
        except Exception as e:
            return jsonify({"error": f"Error reading Excel file: {str(e)}"}), 400

    return jsonify({"error": "Invalid file or no file uploaded."}), 400

@app.route("/edit", methods=["GET"])
def edit():
    filename = request.args.get("filename")
    sheet = request.args.get("sheet")
    file_info = excel_data.get(filename)

    if not file_info or not os.path.exists(file_info["path"]):
        return jsonify({"error": "File not found"}), 404

    df = pd.read_excel(file_info["path"], sheet_name=sheet, dtype=str).fillna("")
    wb = load_workbook(file_info["path"], data_only=True)
    ws = wb[sheet]

    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                options = []
                if dv.formula1.startswith("="):
                    try:
                        ref = dv.formula1.strip("=").replace('$', '')
                        if "!" in ref:
                            sheetname, ref = ref.split("!")
                            target_ws = wb[sheetname]
                        else:
                            target_ws = ws
                        min_col, min_row, max_col, max_row = range_boundaries(ref)
                        for row in target_ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                            for cell in row:
                                if cell.value:
                                    options.append(str(cell.value))
                    except Exception:
                        options = []
                else:
                    options = dv.formula1.strip('"').split(",")

                for cell_range in dv.sqref.ranges:
                    min_col, min_row, max_col, max_row = range_boundaries(str(cell_range))
                    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                        for cell in row:
                            dropdowns[cell.coordinate] = options

    return jsonify({
        "columns": df.columns.tolist(),
        "data": df.to_dict(orient="records"),
        "dropdowns": dropdowns
    })

@app.route("/save", methods=["POST"])
def save():
    data = request.json
    filename = data.get("filename")
    sheet = data.get("sheet")
    edited_data = data.get("data")

    if filename not in excel_data:
        return jsonify({"error": "File not found"}), 404

    filepath = excel_data[filename]["path"]
    df = pd.DataFrame(edited_data)

    wb = load_workbook(filepath)
    ws = wb[sheet]

    # Clear existing data rows (except header)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.value = None

    for r, row_data in enumerate(df.values, start=2):
        for c, value in enumerate(row_data, start=1):
            ws.cell(row=r, column=c, value=value)

    wb.save(filepath)

    # Save second column (edited values) to DB
    try:
        insert_excel_data(df)
    except Exception as e:
        return jsonify({"error": f"DB save failed: {str(e)}"}), 500

    return jsonify({"message": "Saved successfully and logged to DB"})

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
