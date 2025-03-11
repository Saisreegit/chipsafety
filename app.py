from flask import Flask, render_template, request, send_file, jsonify, session
import pandas as pd
import os
from datetime import timedelta

app = Flask(__name__)
app.secret_key = "your_secret_key"
app.config['PERMANENT_SESSION_LIFETIME'] = timedelta(hours=1)  # Keep session alive for 1 hour
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"})
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    session['file_path'] = file_path
    session['filename'] = os.path.splitext(file.filename)[0]  # Store filename without extension
    session['file_type'] = file.filename.rsplit('.', 1)[-1]  # Store file type
    session['modified_filename'] = session['filename']  # Default modified filename
    session['modified_file_type'] = session['file_type']  # Default file type
    session.permanent = True  # Ensure session persists

    return process_file(file_path)

def process_file(file_path, selected_sheet=None):
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            file_type = 'csv'
            sheets = None  
        else:
            excel_file = pd.ExcelFile(file_path)
            sheets = excel_file.sheet_names
            file_type = 'xlsx'
            
            if selected_sheet is None:
                selected_sheet = sheets[0]  

            df = excel_file.parse(selected_sheet)

        df.fillna("", inplace=True)
        return jsonify({
            "data": df.to_dict(orient='records'),
            "file_type": file_type,
            "sheets": sheets,
            "selected_sheet": selected_sheet
        })
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/load_sheet', methods=['POST'])
def load_sheet():
    data = request.json
    file_path = os.path.join(UPLOAD_FOLDER, data.get('filename'))
    selected_sheet = data.get('sheet_name')

    return process_file(file_path, selected_sheet)

@app.route('/update', methods=['POST'])
def update_file():
    try:
        data = request.json['data']
        file_type = request.json.get('file_type', session.get('file_type', 'csv'))
        filename = request.json.get('filename', session.get('filename'))  # User-defined filename
        sheet_name = request.json.get('sheet_name', 'Sheet1')

        if not filename:
            return jsonify({"error": "Filename is required"}), 400

        # Update session variables
        session['modified_filename'] = filename
        session['modified_file_type'] = file_type
        session.permanent = True  # Ensure session persists

        file_path = os.path.join(UPLOAD_FOLDER, f'{filename}.{file_type}')
        df = pd.DataFrame(data)

        if file_type == 'csv':
            df.to_csv(file_path, index=False)
        else:
            with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"✅ File saved successfully: {file_path}")  # Debugging log
        print(f"📝 Stored filename in session: {session.get('modified_filename')}")
        print(f"📂 Available files after update: {os.listdir(UPLOAD_FOLDER)}")
        
        return jsonify({"message": "File updated successfully!", "file_path": file_path, "filename": filename, "file_type": file_type})
    except Exception as e:
        return jsonify({"error": str(e)})

@app.route('/download', methods=['GET'])
def download_file():
    filename = session.get('modified_filename')
    file_type = session.get('modified_file_type', 'csv')
    file_path = os.path.join(UPLOAD_FOLDER, f'{filename}.{file_type}')

    print(f"🔍 Checking for file: {file_path}")
    print(f"📂 Available files: {os.listdir(UPLOAD_FOLDER)}")

    if not os.path.exists(file_path):
        print(f"❌ File not found! Path checked: {file_path}")
        return jsonify({"error": f"File '{filename}.{file_type}' not found. Ensure you have updated before downloading."}), 404

    print(f"📥 Downloading file: {file_path}")
    return send_file(file_path, as_attachment=True, download_name=f'{filename}.{file_type}')

if __name__ == '__main__':
    app.run(debug=True)
