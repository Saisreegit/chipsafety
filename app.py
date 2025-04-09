from flask import Flask, request, send_file, jsonify, render_template
from flask_cors import CORS
import pandas as pd
import os
import tempfile
from datetime import datetime
import mysql.connector
from openpyxl import load_workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils.dataframe import dataframe_to_rows, column_index_from_string
from openpyxl.utils import column_index_from_string

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = tempfile.gettempdir()
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

excel_data = {}  # Cache data in memory

# Environment variables for database connection
DB_HOST = os.getenv("DB_HOST", "your-host")
DB_USER = os.getenv("DB_USER", "your-username")
DB_PASSWORD = os.getenv("DB_PASSWORD", "your-password")
DB_NAME = os.getenv("DB_NAME", "excel_logs")

def get_db_connection():
    return mysql.connector.connect(
        host=DB_HOST,
        user=DB_USER,
        password=DB_PASSWORD,
        database=DB_NAME
    )
def log_to_database(filename, sheet_name):
    print("Logging to DB:", filename, sheet_name)
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS file_logs (
                id INT AUTO_INCREMENT PRIMARY KEY,
                filename VARCHAR(255),
                sheet_name VARCHAR(255),
                status VARCHAR(50),
                timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)
        cursor.execute("INSERT INTO file_logs (filename, sheet_name, status) VALUES (%s, %s, %s)",
                       (filename, sheet_name, 'saved'))
        conn.commit()
        cursor.close()
        conn.close()
        print("✅ DB log inserted")
    except Exception as e:
        print("DB Log Error:", e)
        
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

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    original_headers = [cell.value for cell in ws[1]]
    df = pd.DataFrame(edited_data)[original_headers]
    df.columns = [col.strip() if isinstance(col, str) else col for col in df.columns]
    df.dropna(how='all', inplace=True)
    df = df[[col for col in original_headers if col in df.columns]]

    for i, row in enumerate(dataframe_to_rows(df, index=False, header=False), start=2):
        for j, val in enumerate(row, start=1):
            ws.cell(row=i, column=j).value = val

    wb.save(filepath)
    
    #Log to AWS RDS
    log_to_database(filename, sheet_name)

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
