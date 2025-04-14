from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import os
import tempfile
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

excel_data = {}  # Cache for uploaded Excel files

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    file = request.files["file"]
    if file and file.filename.endswith((".xlsx", ".xls")):
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
        file.save(filepath)
        excel_data[file.filename] = {"path": filepath}
        xl = pd.ExcelFile(filepath)
        return jsonify({"message": "Uploaded", "filename": file.filename, "sheets": xl.sheet_names})
    return jsonify({"error": "Invalid file format"}), 400

@app.route("/edit", methods=["GET"])
def edit():
    filename = request.args.get("filename")
    sheet = request.args.get("sheet")
    file_info = excel_data.get(filename)

    if not file_info or not os.path.exists(file_info["path"]):
        return jsonify({"error": "File not found"}), 404

    df = pd.read_excel(file_info["path"], sheet_name=sheet, dtype=str).fillna("")
    wb = load_workbook(file_info["path"])
    ws = wb[sheet]

    dropdowns = {}
    if ws.data_validations:
        for dv in ws.data_validations.dataValidation:
            if dv.type == "list" and dv.formula1:
                options = dv.formula1.strip('"').split(",")
                for cell_range in dv.sqref.ranges:
                    min_col, min_row, max_col, max_row = range_boundaries(str(cell_range))
                    for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
                        for cell in row:
                            coord = cell.coordinate
                            dropdowns[coord] = options

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

    with pd.ExcelWriter(filepath, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet)

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
