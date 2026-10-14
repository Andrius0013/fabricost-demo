#!/usr/bin/env python3
from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import shutil
import re
from import_project import process_files
import pandas as pd
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 60 * 1024 * 1024

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return jsonify({"error": "Nėra failų"}), 400

    files = request.files.getlist('files')
    if not files or all(not f.filename for f in files):
        return jsonify({"error": "Nepasirinkti failai"}), 400

    temp_dir = tempfile.mkdtemp()
    paths = []
    try:
        for f in files:
            if f.filename:
                safe = re.sub(r'[^\w\.\-]', '_', f.filename)
                p = os.path.join(temp_dir, safe)
                f.save(p)
                paths.append(p)

        result = process_files(paths)
        app.config['LAST_DATA'] = result
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": f"Apdorojimo klaida: {str(e)[:180]}"}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

@app.route('/export_excel')
def export_excel():
    data = app.config.get('LAST_DATA')
    if not data or not data.get('medžiagos'):
        return jsonify({"error": "Nėra duomenų eksportui"}), 400
    df = pd.DataFrame(data['medžiagos'])
    out = BytesIO()
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        df.to_excel(w, index=False, sheet_name='Medžiagų specifikacija')
    out.seek(0)
    return send_file(out, download_name=f"specifikacija_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx", as_attachment=True)

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
