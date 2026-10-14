#!/usr/bin/env python3
from flask import Flask, render_template, request, jsonify, send_file
import os
import tempfile
import shutil
from pathlib import Path
from import_project import process_file, parse_for_materials
import pandas as pd
from io import BytesIO
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 300 * 1024 * 1024

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return jsonify({'error': 'Nėra failų'}), 400
    files = request.files.getlist('files')
    if not files or not files[0].filename:
        return jsonify({'error': 'Nepasirinkti failai'}), 400

    temp_dir = tempfile.mkdtemp()
    try:
        saved = []
        for f in files:
            if f.filename:
                p = os.path.join(temp_dir, Path(f.filename).name)
                f.save(p)
                saved.append(p)
        all_results = []
        for fp in saved:
            all_results.extend(process_file(fp))
        parsed = parse_for_materials(all_results)
        clean = {
            'projektas': parsed.get('projektas', 'Nenustatytas'),
            'medžiagos': parsed.get('medžiagos', []),
            'operacijos': parsed.get('operacijos', []),
            'failai': [Path(f).name for f in saved],
            'total': len(parsed.get('medžiagos', []))
        }
        # Store for export (simple in-memory for demo)
        app.config['LAST_DATA'] = clean
        return jsonify(clean)
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)

@app.route('/export_excel')
def export_excel():
    data = app.config.get('LAST_DATA')
    if not data or not data.get('medžiagos'):
        return jsonify({'error': 'Nėra duomenų eksportui. Pirmiausia įkelkite projektą.'}), 400
    df = pd.DataFrame(data['medžiagos'])
    # Clean columns
    cols = ['pozicija', 'profilis', 'medžiaga', 'ilgis', 'kiekis', 'masė', 'dažomas_plotas', 'kaina', 'šaltinis', 'tipas', 'needs_review']
    df = df.reindex(columns=[c for c in cols if c in df.columns])
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Medžiagų specifikacija')
        # Summary sheet
        summary = pd.DataFrame({
            'Rodiklis': ['Projekto pavadinimas', 'Įrašų skaičius', 'Operacijos', 'Generavimo data'],
            'Reikšmė': [data.get('projektas'), len(data['medžiagos']), ', '.join(data.get('operacijos', [])), datetime.now().strftime('%Y-%m-%d %H:%M')]
        })
        summary.to_excel(writer, index=False, sheet_name='Santrauka')
    output.seek(0)
    return send_file(output, download_name=f"Medziagu_specifikacija_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                     as_attachment=True, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
